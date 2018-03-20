// Office deps
#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.Word"
#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Excel\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.Excel"
#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.PowerPoint\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.PowerPoint"
#I @"C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c"
#r "office"


#I @"..\packages\ExcelProvider.0.8.2\lib"
#r "ExcelProvider.dll"
open FSharp.ExcelProvider


#I @"..\packages\Newtonsoft.Json.10.0.3\lib\net45"
#r "Newtonsoft.Json"
open Newtonsoft.Json

// Need FAKE for @"DocMake\Base\Common.fs"
#I @"..\packages\FAKE.5.0.0-beta005\tools"
#r @"..\packages\FAKE.5.0.0-beta005\tools\FakeLib.dll"


#load @"DocMake\Base\Common.fs"
#load @"DocMake\Base\Office.fs"
#load @"DocMake\Base\Json.fs"
#load @"DocMake\Base\GENHelper.fs"
open DocMake.Base.Common
open DocMake.Base.GENHelper

#load @"DocMake\Tasks\XlsFindReplace.fs"
open DocMake.Tasks.XlsFindReplace

/// This is a one-to-many build, so we don't use FAKE directly, we just use it as a library.

type InputTable = 
    ExcelFile< @"G:\work\Projects\rtu\cema-docs\Site list Yr 3.xlsx",
               SheetName = "Sheet1",
               ForceString = false >

type InputRow = InputTable.Row

let inputTableDict : GetRowsDict<InputTable, InputRow> = 
    { GetRows     = fun imports -> imports.Data 
      NotNullProc = fun row -> match row.GetValue(1) with | null -> false | _ -> true }


let getSiteRows () : InputRow list = 
    excelTableGetRows inputTableDict (new InputTable()) 


//let makeDict1 (row:InputRow) : FindReplaceDict = 
//    Map.ofList 
//        <|  [ "#SITENAME",          row.``RTU - Amp 6, Yr 3 jobs``
//            ; "#WHAT" ,             row.What
//            ]


let main () : unit = 
    let siteList = getSiteRows ()
    printfn "%i sites for output..." siteList.Length
    // Batch file
    //siteList
    //    |> List.map (fun (row:SiteRow) -> row.Name) 
    //    |> generateBatchFile batchConfig 

