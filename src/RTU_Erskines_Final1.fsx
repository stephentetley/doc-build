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

// Need FAKE for @"DocMake\Base\Common.fs" and (@@)
#I @"..\packages\FAKE.5.0.0-beta005\tools"
#r @"..\packages\FAKE.5.0.0-beta005\tools\FakeLib.dll"
open Fake
open Fake.Core
open Fake.Core.Environment
open Fake.Core.Globbing.Operators
open Fake.Core.TargetOperators

#load @"DocMake\Base\Common.fs"
#load @"DocMake\Base\OfficeUtils.fs"
#load @"DocMake\Base\JsonUtils.fs"
#load @"DocMake\Base\GENHelper.fs"
#load @"DocMake\Base\FakeExtras.fs"
open DocMake.Base.Common
open DocMake.Base.GENHelper
open DocMake.Base.FakeExtras

#load @"DocMake\Tasks\DocToPdf.fs"
open DocMake.Tasks.DocToPdf

/// This is a one-to-one build, so we don't use FAKE directly, we just use it as a library.

let _inputRoot = @"G:\work\Projects\rtu\final-docs\Erskines-Batch01-input"
let _outputDir = @"G:\work\Projects\rtu\final-docs\erskine-output\batch01"

let generate1 (dir:string) : unit = 
    match tryFindExactlyOneMatchingFile "*Erskine*.doc*" dir with
    | None -> printfn "BAD:  %s" dir
    | Some ans ->
        let name1 = System.IO.FileInfo(ans).Directory.Name
        printfn "Processing %s ..." name1
        let pdfName = _outputDir @@ (sprintf "%s Erskine Battery Asset Replacement.pdf" name1)
        DocToPdf (fun p -> 
            { p with 
                InputFile = ans
                OutputFile = Some <| pdfName 
            })

let main () = 
    let childDirs = System.IO.Directory.GetDirectories(_inputRoot) |> Array.toList
    List.iter generate1 childDirs


