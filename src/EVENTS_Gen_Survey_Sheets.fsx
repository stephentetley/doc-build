// Office deps
#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.Word"
#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Excel\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.Excel"
#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.PowerPoint\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.PowerPoint"
#I @"C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c"
#r "office"


#I @"..\packages\Newtonsoft.Json.10.0.3\lib\net45"
#r "Newtonsoft.Json"
open Newtonsoft.Json


#I @"..\packages\ExcelProvider.0.8.2\lib"
#r "ExcelProvider.dll"
open FSharp.ExcelProvider


#I @"..\packages\FAKE.5.0.0-beta005\tools"
#r @"..\packages\FAKE.5.0.0-beta005\tools\FakeLib.dll"
open Fake
//open Fake.Core
//open Fake.Core.Environment
//open Fake.Core.Globbing.Operators
//open Fake.Core.TargetOperators

open Microsoft.Office.Interop
open System.Collections.Generic

#load @"DocMake\Base\Common.fs"
#load @"DocMake\Base\OfficeUtils.fs"
#load @"DocMake\Base\JsonUtils.fs"
#load @"DocMake\Base\GENHelper.fs"
#load @"DocMake\Tasks\DocFindReplace.fs"
open DocMake.Base.Common
open DocMake.Base.GENHelper
open DocMake.Tasks.DocFindReplace


/// This is a one-to-many build (one site list, many docs), so 
// we don't use FAKE directly, we just use it as a library.


let _templateRoot   = @"G:\work\Projects\events2\gen-surveys-risks\__Templates"
let _outputRoot     = @"G:\work\Projects\events2\gen-surveys-risks\output"


type SiteTable = 
    ExcelFile< @"G:\work\Projects\events2\EDM2 Site-List SK.xlsx",
               SheetName = "SITE_LIST",
               ForceString = true >

type SiteRow = SiteTable.Row

let siteTableDict : GetRowsDict<SiteTable, SiteRow> = 
    { GetRows     = fun imports -> imports.Data 
      NotNullProc = fun row -> match row.GetValue(0) with | null -> false | _ -> true }



let filterByWorkGroup (workGroup:string) (source:SiteRow list) : SiteRow list = 
    let testRow (row:SiteRow) : bool = 
        match row.``Work Group`` with
        | null -> false
        | ans -> ans = workGroup
    List.filter testRow source

let getSiteRows (workGroup:string) : SiteRow list = 
    excelTableGetRows siteTableDict (new SiteTable()) |> filterByWorkGroup workGroup



let makeTopFolder (batchName:string) : unit = 
    maybeCreateDirectory <| _outputRoot @@ batchName

let makeSiteFolder (batchName:string) (siteName:string) : unit = 
    let cleanName = safeName siteName
    maybeCreateDirectory <| _outputRoot @@ batchName @@ cleanName

let makeSurveyName (siteName:string) (dischargeName:string) : string = 
    sprintf "%s %s Survey.docx" (safeName siteName) (safeName dischargeName)


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

let makeSiteProps (row:SiteRow) : SiteProps = 
    { SiteName = row.``Site Common Name``
      SiteUid = row.``SAI Number`` 
      SiteAddress = row.``Site Address``
      SiteNGR = row.``Site Grid Ref``
      OperationalStatus = row.``Operational Status``
      OperationalResponsibility = row.``Operational Status``
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


let genHazardSheet (workGroup:string)  (site:Site) : unit =
    let template = _templateRoot @@ "TEMPLATE Hazard Identification Check List.docx"
    let cleanName = safeName site.SiteProps.SiteName
    let path1 = _outputRoot @@ workGroup @@ cleanName
    let outPath = path1 @@  (sprintf "%s Hazard Identification Check List.docx" cleanName)    
    DocFindReplace (fun p -> 
        { p with 
            TemplateFile = template
            OutputFile = outPath
            Matches = makeHazardsSearches site.SiteProps 
        }) 


let genSurvey (app:Word.Application) (workGroup:string)  (siteProps:SiteProps) (discharge:Discharge) : unit =
    let template = _templateRoot @@ "TEMPLATE EDM2 Survey 2018-04-24.docx"
    let path1 = _outputRoot @@ safeName workGroup @@ safeName siteProps.SiteName
    let file1 = makeSurveyName siteProps.SiteName discharge.DischargeName
    let outPath = path1 @@ file1
    BatchDocFindReplace app (fun p -> 
        { p with 
            TemplateFile = template
            OutputFile = outPath
            Matches  = makeSurveySearches siteProps discharge
        }) 

// Generating all takes too long just generate a batch.

// TODO ["Harrogate NN", "Leeds", "Sheffield", "Scarborough"]



let main (workGroup:string) : unit = 
    let siteList = buildSites <| getSiteRows workGroup 
    let todoCount = List.length siteList
    let safeBatchName = safeName workGroup
    let app = new Word.ApplicationClass (Visible = true)

    let proc1 (ix:int) (site:Site) = 
        if ix >= 29 then 
            printfn "Generating %i of %i: %s ..." (ix+1) todoCount site.SiteProps.SiteName
            makeSiteFolder safeBatchName site.SiteProps.SiteName
            List.iter (genSurvey app safeBatchName site.SiteProps) site.Discharges
            genHazardSheet safeBatchName site
        else 
            printfn "Skip %i" ix
    // actions...
    makeTopFolder safeBatchName
    siteList |> List.iteri proc1
    
    // finalize...
    app.Quit ()