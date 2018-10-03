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

#I @"..\packages\ExcelProvider.1.0.1\lib\net45"
#r "ExcelProvider.Runtime.dll"

#I @"..\packages\ExcelProvider.1.0.1\typeproviders\fsharp41\net45"
#r "ExcelDataReader.DataSet.dll"
#r "ExcelDataReader.dll"
#r "ExcelProvider.DesignTime.dll"
open FSharp.Interop.Excel





#load "..\src\DocMake\Base\Common.fs"
#load "..\src\DocMake\Base\FakeLike.fs"
#load "..\src\DocMake\Base\OfficeUtils.fs"
#load "..\src\DocMake\Builder\BuildMonad.fs"
#load "..\src\DocMake\Builder\Document.fs"
#load "..\src\DocMake\Builder\Basis.fs"
#load "..\src\DocMake\Tasks\XlsFindReplace.fs"
open DocMake.Base.Common
open DocMake.Base.OfficeUtils
open DocMake.Builder.BuildMonad
open DocMake.Builder.Document
open DocMake.Builder.Basis
open DocMake.Tasks


type InputTable = 
    ExcelFile< @"G:\work\Projects\rtu\cema-docs\Site list Yr 3.xlsx",
               SheetName = "Sheet1",
               ForceString = false >

type InputRow = InputTable.Row

let inputTableDict : ExcelProviderHelperDict<InputTable, InputRow> = 
    { GetRows     = fun imports -> imports.Data 
      NotNullProc = fun row -> match row.GetValue(1) with | null -> false | _ -> true }


let getSiteRows () : InputRow list = 
    excelTableGetRows inputTableDict (new InputTable()) |> Seq.toList


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
let xlsFindReplace = api.XlsFindReplace
let getTemplate = api.GetTemplateXls



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
          PrintQuality = PrintQuality.PqScreen
          PdfQuality = PdfPrintQuality.PdfScreen }

    let excelApp = initExcel ()
    let excelKill = fun (app:Excel.Application) -> finalizeExcel app

    consoleRun env excelApp excelKill (buildScript ())



