// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


// Office deps
#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.Word"
#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Excel\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.Excel"
#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.PowerPoint\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.PowerPoint"
#I @"C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c"
#r "office"


open Microsoft.Office.Interop


#I @"..\packages\Newtonsoft.Json.11.0.2\lib\net45"
#r "Newtonsoft.Json"
open Newtonsoft.Json


#I @"..\packages\ExcelProvider.0.8.2\lib"
#r "ExcelProvider.dll"
open FSharp.ExcelProvider



#load @"DocMake\Base\Common.fs"
#load @"DocMake\Base\FakeLike.fs"
#load @"DocMake\Base\ExcelProviderHelper.fs"
#load @"DocMake\Base\OfficeUtils.fs"
#load @"DocMake\Base\SimpleDocOutput.fs"
#load @"DocMake\Builder\BuildMonad.fs"
#load @"DocMake\Builder\Document.fs"
#load @"DocMake\Builder\Basis.fs"
#load @"DocMake\Tasks\DocFindReplace.fs"
#load @"DocMake\WordBuilder.fs"
open DocMake.Base.Common
open DocMake.Base.FakeLike
open DocMake.Base.ExcelProviderHelper
open DocMake.Builder.BuildMonad
open DocMake.Builder.Document
open DocMake.Builder.Basis
open DocMake.WordBuilder


/// This is a one-to-many build - one site list, many docs.


let _templateRoot   = @"G:\work\Projects\events2\gen-surveys-risks\__Templates"
let _outputRoot     = @"G:\work\Projects\events2\gen-surveys-risks\output"



type SiteTable = 
    ExcelFile< @"G:\work\Projects\events2\EDM2 Site-List SK.xlsx",
               SheetName = "SITE_LIST",
               ForceString = true >

type SiteRow = SiteTable.Row


let readSiteRows () : SiteRow list = 
    let helper = 
        { new IExcelProviderHelper<SiteTable,SiteRow>
          with member this.ReadTableRows table = table.Data 
               member this.IsBlankRow row = match row.GetValue(0) with null -> true | _ -> false }
    excelReadRowsAsList helper (new SiteTable())


let filterBySurveyBatch (batch:string) (source:SiteRow list) : SiteRow list = 
    let testRow (row:SiteRow) : bool = 
        match row.``Survey Batch`` with
        | null -> false
        | ans -> ans = batch
    List.filter testRow source

let getSiteRows (surveyBatch:string) : SiteRow list = 
    readSiteRows ()
        |> filterBySurveyBatch surveyBatch



let makeTopFolder (batchName:string) : unit = 
    maybeCreateDirectory <| (_outputRoot </> batchName)

let makeSiteFolder (batchName:string) (siteName:string) : unit = 
    let cleanName = safeName siteName
    maybeCreateDirectory <| (_outputRoot </> batchName </> cleanName)

let makeSurveyName (siteName:string) (dischargeName:string) : string = 
    sprintf "%s %s Survey.docx" (safeName siteName) (safeName dischargeName)

let makeUSRecalcName (siteName:string) (dischargeName:string) : string = 
    sprintf "%s %s US Calibration.docx" (safeName siteName) (safeName dischargeName)

type SiteProps = 
    { SiteName: string
      SiteUid: string
      SiteAddress: string
      SiteNGR: string
      OperationalStatus: string
      OperationalResponsibility: string
      AssetType: string }

type Discharge = 
    { DischargeName: string
      OutfallNGR: string
      Watercourse: string }


type Site = 
    { SiteProps: SiteProps
      Discharges: Discharge  list }

type InterimSites = Map<string,Site>



let makeSurveySearches (site:SiteProps) (discharge:Discharge) : SearchList = 
    [ "#SITENAME",          site.SiteName
    ; "#SAINUMBER" ,        site.SiteUid
    ; "#SITEADDRESS",       site.SiteAddress
    ; "#OPERSTATUS",        site.OperationalStatus
    ; "#SITEGRIDREF",       site.SiteNGR
    ; "#ASSETTYPE",         site.AssetType
    ; "#DISCHARGENAME",     discharge.DischargeName
    ; "#OPERNAME",          site.OperationalResponsibility
    ; "#OUTFALLGRIDREF",    discharge.OutfallNGR
    ; "#RECWATERWOURSE",    discharge.Watercourse
    ]



let makeHazardsSearches (site:SiteProps) : SearchList = 
    [ "#SITENAME",          site.SiteName   
    ; "#SAINUMBER" ,        site.SiteUid

    ]


let makeUSRecalcSearches (site:SiteProps) (discharge:Discharge) : SearchList = 
    [ "#SITENAME",                  site.SiteName   
    ; "#SAINUMBER" ,                site.SiteUid
    ; "#DISCHARGENAME",             discharge.DischargeName
    ]

let makeSiteProps (row:SiteRow) : SiteProps = 
    { SiteName = row.``Site Common Name``
      SiteUid = row.``SAI Number`` 
      SiteAddress = row.``Site Address``
      SiteNGR = row.``Site Grid Ref``
      OperationalStatus = row.``Operational Status``
      OperationalResponsibility = row.``Operational Responsibility``
      AssetType = row.``Site Type`` }

let makeDischargeProps (row:SiteRow) : Discharge = 
    { DischargeName = row.``Discharge Name``
      OutfallNGR = row.``Outfall Grid Ref (from IW sheet, may lack precision)``
      Watercourse = row.``Receiving Watercourse`` }


// Tree Build API
// Exists, Insert, Update (which is just add child)

let exists (key:string) (forest:InterimSites) : bool = forest.ContainsKey key

let insert (site:Site) (forest:InterimSites) : InterimSites = 
    forest.Add(site.SiteProps.SiteName, site)

let addChild (siteName:string) (discharge:Discharge) (forest:InterimSites) : InterimSites = 
    match forest.TryFind siteName with
    | Some site -> 
        let site2 = { site with Discharges = site.Discharges @ [discharge] }
        forest.Add(siteName, site2)
    | None -> forest

let buildSites1 (row:SiteRow) (forest:InterimSites) : InterimSites =
    let siteName = row.``Site Common Name``
    let siteProps = makeSiteProps row
    let discharge = makeDischargeProps row
    if exists siteName forest then
        addChild siteName discharge forest
    else
        insert {SiteProps = siteProps; Discharges = [discharge]} forest

let buildSites (rows: SiteRow list) : Site list = 
    let interim = List.fold (fun st row -> buildSites1 row st) Map.empty rows 
    interim |> Map.toList |> List.map snd


// ********************************
// Build script



let genHazardSheet (workGroup:string)  (site:Site) : WordBuild<WordDoc> =
    buildMonad { 
        let templatePath = _templateRoot </> "TEMPLATE Hazard Identification Check List.docx"
        let cleanName = safeName site.SiteProps.SiteName
        let subPath = safeName workGroup </> cleanName
        let outName = sprintf "%s Hazard Identification Check List.docx" cleanName   
        let matches = makeHazardsSearches site.SiteProps 
        let! d1 = 
            localSubDirectory subPath <| 
                buildMonad { 
                    let! template = getTemplateDoc templatePath
                    let! d1 = docFindReplace matches template >>= renameTo outName 
                    return d1
                }
        return d1
    }

let genSurvey (workGroup:string)  (siteProps:SiteProps) (discharge:Discharge) : WordBuild<WordDoc> = 
    buildMonad { 
        let surveyTemplate = _templateRoot </> "TEMPLATE EDM2 Survey 2018-05-31.docx"
        let subPath = safeName workGroup </> safeName siteProps.SiteName
        let outName = makeSurveyName siteProps.SiteName discharge.DischargeName
        let matches = makeSurveySearches siteProps discharge
        let! d1 = 
            localSubDirectory subPath <| 
                buildMonad { 
                    let! template = getTemplateDoc surveyTemplate
                    let! d1 = docFindReplace matches template >>= renameTo outName 
                    return d1
                }
        return d1
    }

    
let genUSRecalc (workGroup:string)  (siteProps:SiteProps) (discharge:Discharge) : WordBuild<WordDoc> = 
    buildMonad { 
        let surveyTemplate = _templateRoot </> "TEMPLATE US Calibration Record.docx"
        let subPath = safeName workGroup </> safeName siteProps.SiteName
        let outName = makeUSRecalcName siteProps.SiteName discharge.DischargeName
        let matches = makeUSRecalcSearches siteProps discharge
        let! d1 = 
            localSubDirectory subPath <| 
                buildMonad { 
                    let! template = getTemplateDoc surveyTemplate
                    let! d1 = docFindReplace matches template >>= renameTo outName 
                    return d1
                }
        return d1
    }
// Generating all takes too long just generate a batch.

// TODO SHEFFIELD , YORK

let buildScript (surveyBatch:string) (makeHazards:bool) : WordBuild<unit> = 
    let siteList = buildSites <| getSiteRows surveyBatch 
    let todoCount = List.length siteList
    let safeBatchName = safeName surveyBatch

    let proc1 (ix:int) (site:Site) : WordBuild<unit> = 
        buildMonad { 
            do printfn "Generating %i of %i: %s ..." (ix+1) todoCount site.SiteProps.SiteName
            // do makeSiteFolder safeBatchName site.SiteProps.SiteName
            
            // Survey
            // do! forMz site.Discharges (genSurvey safeBatchName site.SiteProps) 
            if makeHazards then 
                do! genHazardSheet safeBatchName site |>> ignore
            else return ()
            // US Calibration
            do! forMz site.Discharges (genUSRecalc safeBatchName site.SiteProps) 
            return ()
        }

    foriMz siteList proc1
                               


let main (surveyBatch:string) (makeHazards:bool) : unit = 
    let env = 
        { WorkingDirectory = _outputRoot
          PrintQuality = PrintQuality.PqScreen
          PdfQuality = PdfPrintQuality.PdfScreen }

    runWordBuild env (buildScript surveyBatch makeHazards)

let allBatches () : Set<string> = 
    let allRows = readSiteRows ()
    let add1 = 
        fun ac (row:SiteRow) -> 
            match row.``Survey Batch`` with
            | null -> ac
            | name -> Set.add name ac
    List.fold add1 Set.empty allRows

let main2 () = 
    allBatches () 
        |> Set.iter (fun name -> main name false)
