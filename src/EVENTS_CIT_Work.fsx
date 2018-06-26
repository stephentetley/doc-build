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



#I @"..\packages\ExcelProvider.0.8.2\lib"
#r "ExcelProvider.dll"
open FSharp.ExcelProvider


open System.IO


#load @"DocMake\Base\Common.fs"
#load @"DocMake\Base\FakeLike.fs"
#load @"DocMake\Base\OfficeUtils.fs"
#load @"DocMake\Base\SimpleDocOutput.fs"
#load @"DocMake\Builder\BuildMonad.fs"
#load @"DocMake\Builder\Basis.fs"
#load @"DocMake\Builder\WordHooks.fs"
open DocMake.Base.Common
open DocMake.Base.FakeLike
open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis
open DocMake.Builder.WordHooks

#load @"DocMake\Tasks\DocFindReplace.fs"
open DocMake.Tasks


// Simple find-and-replace (mail merge-like).
// Generate multiple outputs in a single folder.


let _templateRoot       = @"G:\work\Projects\events2\gen-cit-sheets\__Templates"
let _outputDirectory    = @"G:\work\Projects\events2\gen-cit-sheets\output"

let _surveyTemplate = _templateRoot </> "TEMPLATE Scope of Works.docx"

type SiteTable = 
    ExcelFile< @"G:\work\Projects\events2\EDM2 Site-List SK.xlsx",
               SheetName = "SITE_LIST",
               ForceString = true >

type SiteRow = SiteTable.Row

let siteTableDict : ExcelProviderHelperDict<SiteTable, SiteRow> = 
    { GetRows     = fun imports -> imports.Data 
      NotNullProc = fun row -> match row.GetValue(0) with | null -> false | _ -> true }

let filterBySurveyComplete (source:SiteRow list) : SiteRow list = 
    let testRow (row:SiteRow) : bool = 
        match row.``Survey Completed (initials - date)`` with
        | null -> false
        | ans -> ans <> ""
    List.filter (not << testRow) source

let getSiteRows () : SiteRow list = 
    excelTableGetRows siteTableDict (new SiteTable()) 
        |> Seq.toList 
        |> filterBySurveyComplete


let makeMatches (row:SiteRow) : SearchList = 
    let dNow = System.DateTime.Now
    [ "#SITENAME",          row.``Site Common Name``
    ; "#TODAY",             dNow.ToString "dd/MM/yyyy"
    ; "#SAINUMBER" ,        row.``SAI Number``
    ; "#SITEADDRESS",       row.``Site Address``
    ; "#OPERSTATUS",        row.``Operational Status``
    ; "#SITEGRIDREF",       row.``Site Grid Ref``
    ; "#ASSETTYPE",         row.``Site Type``
    ; "#OPERNAME",          row.``Operational Responsibility``
    ; "#DISCHARGENAME",     row.``Discharge Name``
    ; "#OPERNAME",          row.``Operational Responsibility``
    ; "#OUTFALLGRIDREF",    row.``Outfall Grid Ref (from IW sheet, may lack precision)``
    ; "#RECWATERCOURSE",    row.``Receiving Watercourse``
    ]


// ********************************
// Build script


type EventsRes = Word.Application

type EventsBuild<'a> = BuildMonad<EventsRes,'a>

// Just need the DocFindReplace API...
let api = DocFindReplace.makeAPI (fun app -> app)
let docFindReplace = api.docFindReplace
let getTemplate = api.getTemplateDoc


let scopeOfWorks (row:SiteRow) : EventsBuild<WordDoc> = 
    buildMonad { 
        let name1 = 
            let s = safeName row.``Site Common Name``
            if s.Length > 40 then
                s.[0..39]
            else
                s

        let docName = sprintf "%s Scope of Works.docx" name1
        let matches = makeMatches row
        let! template = getTemplate _surveyTemplate
        let! d1 = docFindReplace matches template >>= renameTo docName
        return d1 } 

let buildScript () : EventsBuild<unit> = 
    let siteList = getSiteRows () |> List.sortBy (fun r -> r.``Site Common Name``)
    let count = siteList.Length
    foriMz siteList <| fun ix row -> 
        if ix > 0 then
            printfn "Site %i of %i:" (ix+1) count
            fmapM ignore <| scopeOfWorks row
        else
            breturn ()
        

let main () : unit = 
    let env = 
        { WorkingDirectory = _outputDirectory
          PrintQuality = DocMakePrintQuality.PqScreen
          PdfQuality = PdfPrintSetting.PdfScreen }
    let hooks:BuilderHooks<Word.Application> = wordBuilderHook
    consoleRun env hooks (buildScript ())