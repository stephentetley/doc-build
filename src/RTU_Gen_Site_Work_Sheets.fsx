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

// Use FSharp.Data for CSV output (Proprietry.fs)
#I @"..\packages\FSharp.Data.3.0.0-beta3\lib\net45"
#r @"FSharp.Data.dll"

// Use ExcelProvider to read SAI numbers spreadsheet (Proprietry.fs)
#I @"..\packages\ExcelProvider.0.8.2\lib"
#r "ExcelProvider.dll"
open FSharp.ExcelProvider

#load @"DocMake\Base\Common.fs"
#load @"DocMake\Base\FakeLike.fs"
#load @"DocMake\Base\OfficeUtils.fs"
#load @"DocMake\Builder\BuildMonad.fs"
#load @"DocMake\Builder\Document.fs"
#load @"DocMake\Builder\Basis.fs"
#load @"DocMake\Tasks\DocFindReplace.fs"
#load @"Proprietry.fs"
open DocMake.Base.Common
open DocMake.Base.FakeLike
open DocMake.Base.OfficeUtils
open DocMake.Builder.BuildMonad
open DocMake.Builder.Document
open DocMake.Builder.Basis
open DocMake.Tasks
open Proprietry

/// TODO - update to work with BuildMonad

/// This is a one-to-many build (one site list generates many docs), 
/// so we don't use FAKE directly, we just use it as a library.


let _templateRoot       = @"G:\work\Projects\rtu\site-docs\__Templates"
let _outputRoot         = @"G:\work\Projects\rtu\site-docs\output\Batch01"


type SiteTable = 
    ExcelFile< @"G:\work\Projects\rtu\site-docs\year4_sitelist.xlsx",
               SheetName = "SITE_LIST",
               ForceString = false >

type SiteRow = SiteTable.Row

let siteTableDict : ExcelProviderHelperDict<SiteTable, SiteRow> = 
    { GetRows     = fun imports -> imports.Data 
      NotNullProc = fun row -> match row.GetValue(0) with | null -> false | _ -> true }



let getSiteRows () : SiteRow list = 
    excelTableGetRows siteTableDict (new SiteTable()) |> Seq.toList




let makeSurveySearches (row:SiteRow) : SearchList = 
    [ "#SITENAME",          row.``Common Name``
    ; "#OSNAME" ,           row.``OS Name``
    ; "#OSADDR",            row.``Os Addr``
    ]


type WordRes = Word.Application

type WordBuild<'a> = BuildMonad<WordRes,'a>

// Just need the DocFindReplace API...
let api = DocFindReplace.makeAPI (fun app -> app)
let docFindReplace = api.DocFindReplace
let getTemplate = api.GetTemplateDoc



let genSurvey (row:SiteRow) : WordBuild<WordDoc> =
    buildMonad { 
        let! template = getTemplate (_templateRoot </> "TEMPLATE RTU Site Works Record.docx")
        let cleanName = underscoreName row.``Common Name``
        let surveyName = sprintf "%s Site Works Record.docx" cleanName
        let! d1 = docFindReplace (makeSurveySearches row) template >>= renameTo surveyName
        return d1
    }


let makeHazardsSearches (row:SiteRow) : SearchList = 
    [ "#SAINUMBER" ,        row.Uid
    ; "#SITENAME",          row.``Common Name``
    ]

let genHazardSheet (row:SiteRow) : WordBuild<WordDoc> =
    buildMonad { 
        let! template = getTemplate (_templateRoot </> "TEMPLATE Hazard Identification Check List.docx")
        let cleanName = underscoreName row.``Common Name``
        let formName = sprintf "%s Hazard Identification Check List.docx" cleanName
        let! d1 = docFindReplace (makeHazardsSearches row) template >>= renameTo formName
        return d1
    }




let main () : unit = 
    let siteList = getSiteRows () 
    let todoCount = List.length siteList

    let procM (ix:int) (row:SiteRow) : WordBuild<unit> = 
        printfn "Generating %i of %i: %s ..." (ix+1) todoCount row.``Common Name``
        localSubDirectory (underscoreName row.``Common Name``) 
            <| buildMonad { 
                    let! _ = genSurvey row
                    let! _ = genHazardSheet row
                    return ()
                }
    
    let env = 
        { WorkingDirectory = _outputRoot
          PrintQuality = PrintQuality.PqScreen
          PdfQuality = PdfPrintQuality.PdfScreen }
    
    let wordApp = initWord ()
    let wordKill = fun (app:Word.Application) -> finalizeWord app
    consoleRun env wordApp wordKill (foriMz siteList procM)