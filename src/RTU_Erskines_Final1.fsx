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


#I @"..\packages\ExcelProvider.0.8.2\lib"
#r "ExcelProvider.dll"
open FSharp.ExcelProvider


#I @"..\packages\Newtonsoft.Json.11.0.2\lib\net45"
#r "Newtonsoft.Json"


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
open Fake.IO.FileSystemOperators

#load @"DocMake\Base\Common.fs"
#load @"DocMake\Base\OfficeUtils.fs"
#load @"DocMake\Base\JsonUtils.fs"
#load @"DocMake\Base\GENHelper.fs"
#load @"DocMake\Base\FakeFake.fs"
#load @"DocMake\Base\FakeExtras.fs"
open DocMake.Base.Common
open DocMake.Base.FakeExtras

#load @"DocMake\Tasks\DocToPdf.fs"
open DocMake.Tasks.DocToPdf

/// TODO change to use Build monad

let _inputRoot = @"G:\work\Projects\rtu\final-docs\input\Erskines_02"
let _outputDir = @"G:\work\Projects\rtu\final-docs\output\erskine-output\batch02"

let generate1 (dir:string) : unit = 
    match tryFindExactlyOneMatchingFile "*Erskine*.doc*" dir with
    | None -> printfn "BAD:  %s" dir
    | Some ans ->
        let name1 = System.IO.FileInfo(ans).Directory.Name
        printfn "Processing %s ..." name1
        let pdfName = _outputDir </> (sprintf "%s Erskine Battery Asset Replacement.pdf" name1)
        printfn "Output file: %s" pdfName
        DocToPdf (fun p -> 
            { p with 
                InputFile = ans
                OutputFile = Some <| pdfName 
            })

let main () = 
    maybeCreateDirectory _outputDir
    let childDirs = System.IO.Directory.GetDirectories(_inputRoot)
    Array.iter generate1 childDirs


