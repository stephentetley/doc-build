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


#I @"..\packages\Newtonsoft.Json.11.0.2\lib\net45"
#r "Newtonsoft.Json"
open Newtonsoft.Json

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
#load @"DocMake\Base\JsonUtils.fs"
#load @"DocMake\Base\GENHelper.fs"
#load @"DocMake\Tasks\XlsFindReplace.fs"
open DocMake.Base.Common
open DocMake.Base.GENHelper
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


let makeSearches (row:InputRow) : SearchList = 
    [ "#SITENAME",          row.``RTU - Amp 6, Yr 3 jobs``
    ; "#WHAT" ,             row.What
    ]

    
let _outputRoot     = @"G:\work\Projects\rtu\cema-docs\output"
let _template       = @"G:\work\Projects\rtu\cema-docs\__Templates\TEMPLATE Customer Information Hazardous Area Sheet.xlsx"


// Note it's ineffienct to repeated create an instance of Excel foreach row.
// Look at providing two routes into XlsFindReplace
let process1  (row:InputRow) : unit =
    let cleanName = safeName row.``RTU - Amp 6, Yr 3 jobs``
    let outPath = _outputRoot @@ (sprintf "%s HS Form 81-1 Hazardous Areas.xlsx" cleanName)
    XlsFindReplace (fun p -> 
        { p with 
            TemplateFile = _template
            OutputFile = outPath
            Matches  = makeSearches row
        }) 

let main () : unit = 
    let siteList = getSiteRows ()
    printfn "%i sites for output..." siteList.Length
    siteList |> List.iter process1

