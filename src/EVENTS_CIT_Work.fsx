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

// FAKE dependencies are getting onorous...
#I @"..\packages\FAKE.5.0.0-rc016.225\tools"
#r @"FakeLib.dll"
#I @"..\packages\Fake.Core.Globbing.5.0.0-beta021\lib\net46"
#r @"Fake.Core.Globbing.dll"
#I @"..\packages\Fake.IO.FileSystem.5.0.0-rc017.237\lib\net46"
#r @"Fake.IO.FileSystem.dll"
#I @"..\packages\Fake.Core.Trace.5.0.0-rc017.237\lib\net46"
#r @"Fake.Core.Trace.dll"
#I @"..\packages\Fake.Core.Process.5.0.0-rc017.237\lib\net46"
#r @"Fake.Core.Process.dll"
open Fake
open Fake.IO.FileSystemOperators



#load @"DocMake\Base\Common.fs"
#load @"DocMake\Base\OfficeUtils.fs"
#load @"DocMake\Base\SimpleDocOutput.fs"
#load @"DocMake\Builder\BuildMonad.fs"
#load @"DocMake\Builder\Basis.fs"
#load @"DocMake\Builder\WordBuilder.fs"
open DocMake.Base.Common
open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis
open DocMake.Builder.WordBuilder

#load @"DocMake\Lib\DocFindReplace.fs"
open DocMake.Lib.DocFindReplace


// Simple find-and-replace (mail merge-like).
// Generate multiple outputs in a single folder.


let _templateRoot       = @"G:\work\Projects\events2\gen-cit-sheets\__Templates"
let _outputDirectory    = @"G:\work\Projects\events2\gen-cit-sheets\output"

let _surveyTemplate = _templateRoot @@ "TEMPLATE Scope of Works.docx"

type SiteTable = 
    ExcelFile< @"G:\work\Projects\events2\EDM2 Site-List SK.xlsx",
               SheetName = "SITE_LIST",
               ForceString = true >

type SiteRow = SiteTable.Row

let siteTableDict : ExcelProviderHelperDict<SiteTable, SiteRow> = 
    { GetRows     = fun imports -> imports.Data 
      NotNullProc = fun row -> match row.GetValue(0) with | null -> false | _ -> true }

let filterBySurveyBatch (batch:string) (source:SiteRow list) : SiteRow list = 
    let testRow (row:SiteRow) : bool = 
        match row.``Survey Batch`` with
        | null -> false
        | ans -> ans = batch
    List.filter testRow source

let getSiteRows (surveyBatch:string) : SiteRow list = 
    excelTableGetRows siteTableDict (new SiteTable()) // |> filterBySurveyBatch surveyBatch


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

let scopeOfWorks (row:SiteRow) : BuildMonad<'res, WordDoc> = 
    execWordBuild ( 
        buildMonad { 
            let docName = sprintf "%s Scope of Works.docx" (safeName row.``Site Common Name``)
            let matches = makeMatches row
            let! template = getTemplate _surveyTemplate
            let! d1 = docFindReplace matches template >>= renameTo docName
            return d1 }) 

let buildScript () : BuildMonad<'res,unit> = 
    let siteList = List.take 5 <|  getSiteRows "" 
    forMz siteList scopeOfWorks

type EventsRes = 
    { WordApp : Word.Application } 




let main () : unit = 
    let env = 
        { WorkingDirectory = _outputDirectory
          PrintQuality = DocMakePrintQuality.PqScreen
          PdfQuality = PdfPrintSetting.PdfScreen }

    consoleRun env (buildScript ()) 