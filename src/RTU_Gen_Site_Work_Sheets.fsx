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


#I @"..\packages\Newtonsoft.Json.11.0.2\lib\net45"
#r "Newtonsoft.Json"
open Newtonsoft.Json


#I @"..\packages\ExcelProvider.0.8.2\lib"
#r "ExcelProvider.dll"
open FSharp.ExcelProvider


#load @"DocMake\Base\Common.fs"
#load @"DocMake\Base\OfficeUtils.fs"
#load @"DocMake\Base\JsonUtils.fs"
#load @"DocMake\Tasks\DocFindReplace.fs"
open DocMake.Base.Common
open DocMake.Tasks.DocFindReplace

/// TODO - update to work with BuildMonad

/// This is a one-to-many build (one site list generates many docs), 
/// so we don't use FAKE directly, we just use it as a library.


let _templateRoot   = @"G:\work\Projects\rtu\site-docs\__Templates"
let _outputRoot     = @"G:\work\Projects\rtu\site-docs\output"


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


let makeSiteFolder (siteName:string) : unit = 
    let cleanName = safeName siteName
    maybeCreateDirectory <| _outputRoot @@  cleanName


let makeSurveySearches (row:SiteRow) : SearchList = 
    [ "#SITENAME",          row.``Common Name``
    ; "#OSNAME" ,           row.``OS Name``
    ; "#OSADDR",            row.``Os Addr``
    ]


let genSurvey (row:SiteRow) : unit =
    let template = _templateRoot @@ "TEMPLATE RTU Site Works Record.docx"
    let cleanName = safeName row.``Common Name``
    let path1 = _outputRoot @@ cleanName
    let outPath = path1 @@ (sprintf "%s Site Works Record.docx" cleanName)
    DocFindReplace (fun p -> 
        { p with 
            TemplateFile = template
            OutputFile = outPath
            Matches  = makeSurveySearches row
        }) 

let makeHazardsSearches (row:SiteRow) : SearchList = 
    [ "#SAINUMBER" ,        row.Uid
    ; "#SITENAME",          row.``Common Name``
    ]

let genHazardSheet (row:SiteRow) : unit =
    let template = _templateRoot @@ "TEMPLATE Hazard Identification Check List.docx"
    let cleanName = safeName row.``Common Name``
    let path1 = _outputRoot @@ cleanName
    let outPath = path1 @@  (sprintf "%s Hazard Identification Check List.docx" cleanName)    
    DocFindReplace (fun p -> 
        { p with 
            TemplateFile = template
            OutputFile = outPath
            Matches = makeHazardsSearches row 
        }) 



let main () : unit = 
    let siteList = getSiteRows () 
    let todoCount = List.length siteList

    let proc1 (ix:int) (row:SiteRow) = 
        printfn "Generating %i of %i: %s ..." (ix+1) todoCount row.``Common Name``
        makeSiteFolder row.``Common Name``
        genSurvey row
        genHazardSheet row
    
    // actions...
    siteList |> List.iteri proc1
