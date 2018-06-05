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



// Need FAKE for @"DocMake\Base\Common.fs" and (@@)
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


#load @"DocMake\Base\Common.fs"
#load @"DocMake\Base\OfficeUtils.fs"
#load @"DocMake\Builder\BuildMonad.fs"
#load @"DocMake\Builder\Basis.fs"
#load @"DocMake\Builder\ExcelBuilder.fs"
open DocMake.Base.Common
open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis
open DocMake.Builder.ExcelBuilder

#load @"DocMake\Lib\XlsFindReplace.fs"
open DocMake.Lib


/// This is a one-to-many build, so we don't use FAKE directly, we just use it as a library.

type InputTable = 
    ExcelFile< @"G:\work\Projects\rtu\cema-docs\Site list Yr 3.xlsx",
               SheetName = "Sheet1",
               ForceString = false >

type InputRow = InputTable.Row

let inputTableDict : ExcelProviderHelperDict<InputTable, InputRow> = 
    { GetRows     = fun imports -> imports.Data 
      NotNullProc = fun row -> match row.GetValue(1) with | null -> false | _ -> true }


let getSiteRows () : InputRow list = 
    excelTableGetRows inputTableDict (new InputTable()) 


let makeSearches (row:InputRow) : SearchList = 
    [ "#SITENAME",          row.``RTU - Amp 6, Yr 3 jobs``
    ; "#WHAT" ,             row.What
    ]

    
let _outputRoot     = @"G:\work\Projects\rtu\cema-docs\output"
let _template       = @"G:\work\Projects\rtu\cema-docs\__Templates\TEMPLATE Customer Information Hazardous Area Sheet.xlsx"



type ExcelRes = Excel.Application

type ExcelBuild<'a> = BuildMonad<ExcelRes,'a>

// Just need the XlsFindReplace API...
let api = XlsFindReplace.makeAPI (fun app -> app)
let xlsFindReplace = api.xlsFindReplace
let getTemplate = api.getTemplate



// Note it's ineffienct to repeated create an instance of Excel foreach row.
// Look at providing two routes into XlsFindReplace
let hazardSheet  (row:InputRow) : ExcelBuild<ExcelDoc> =
    buildMonad { 
        let cleanName = safeName row.``RTU - Amp 6, Yr 3 jobs``
        let outName = sprintf "%s HS Form 81-1 Hazardous Areas.xlsx" cleanName
        let matches = makeSearches row
        let! template = getTemplate _template
        let! d1 = xlsFindReplace matches template >>= renameTo outName
        return d1
    }

let buildScript () : ExcelBuild<unit> = 
    let siteList = getSiteRows ()
    forMz siteList hazardSheet


let main () : unit = 
    let siteList = getSiteRows ()
    printfn "%i sites for output..." siteList.Length

    let env = 
        { WorkingDirectory = _outputRoot
          PrintQuality = DocMakePrintQuality.PqScreen
          PdfQuality = PdfPrintSetting.PdfScreen }
    let hooks:BuilderHooks<Excel.Application> = excelBuilderHook
    consoleRun env hooks (buildScript ())



