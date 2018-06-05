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


#I @"..\packages\Magick.NET-Q8-AnyCPU.7.4.6\lib\net40"
#r @"Magick.NET-Q8-AnyCPU.dll"
open ImageMagick

#I @"..\packages\Newtonsoft.Json.11.0.2\lib\net45"
#r "Newtonsoft.Json"
open Newtonsoft.Json

open System.IO

// FAKE is local to the project file
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
open Fake.Core
open Fake.IO.FileSystemOperators


#load @"DocMake\Base\Common.fs"
#load @"DocMake\Base\String.fs"
#load @"DocMake\Base\FakeExtras.fs"
#load @"DocMake\Base\JsonUtils.fs"
#load @"DocMake\Base\ImageMagickUtils.fs"
#load @"DocMake\Base\JsonUtils.fs"
#load @"DocMake\Base\OfficeUtils.fs"
#load @"DocMake\Base\CopyRename.fs"
#load @"DocMake\Base\SimpleDocOutput.fs"
open DocMake.Base.Common
open DocMake.Base.String
open DocMake.Base.FakeExtras
open DocMake.Base.CopyRename


#load @"DocMake\Tasks\DocFindReplace.fs"
#load @"DocMake\Tasks\DocPhotos.fs"
#load @"DocMake\Tasks\DocToPdf.fs"
#load @"DocMake\Tasks\XlsToPdf.fs"
#load @"DocMake\Tasks\PdfConcat.fs"
open DocMake.Tasks.DocFindReplace
open DocMake.Tasks.DocPhotos
open DocMake.Tasks.DocToPdf
open DocMake.Tasks.XlsToPdf
open DocMake.Tasks.PdfConcat

// NOTE - can generate a batch file to do "many-to-one"
// We usually have many final docs to make (many assets) but the style of 
// Fake is to make one agglomerate out of many parts.
// Generating a batch file that invokes Fake for each asset solves this.

let _filestoreRoot  = @"G:\work\Projects\plcar\final_docs\input\FILEY_STW"
let _outputRoot     = @"G:\work\Projects\plcar\final_docs\output\FILEY_STW"
let _templateRoot   = @"G:\work\Projects\plcar\final_docs\__Templates"
let _jsonRoot       = @"G:\work\Projects\plcar\final_docs\__Json"


let assetName = "FILEY/STW"
let coverMatches : SearchList = 
    [ "#SITENAME", assetName
    ; "#SAINUM", "SAI00179696"
    ; "#PLC", ""
    ]


let cleanName           = safeName assetName
let assetInputDir       = _filestoreRoot @@ cleanName
let assetOutputDir      = _outputRoot


let makeAssetOutputName (fmt:Printf.StringFormat<string->string>) : string = 
    assetOutputDir @@ sprintf fmt cleanName

let makeClean () : unit =  
    if Directory.Exists(assetOutputDir) then 
        Trace.tracefn " --- Clean folder: '%s' ---" assetOutputDir
        Fake.IO.Directory.delete assetOutputDir
    else 
        Trace.tracefn " --- Clean --- : folder does not exist '%s' ---" assetOutputDir


let makeOutputDirectory () : unit =  
    Trace.tracefn " --- Output folder: '%s' ---" assetOutputDir
    maybeCreateDirectory assetOutputDir 



let makeCover () : unit = 
    try
        let template = _templateRoot @@ "TEMPLATE PLC Cover Sheet.docx"
        let docName = makeAssetOutputName "%s cover-sheet.docx"
        let pdfName = pathChangeExtension docName "pdf"

        DocFindReplace (fun p -> 
            { p with 
                TemplateFile = template
                OutputFile = docName
                Matches  = coverMatches 
            }) 
        DocToPdf (fun p -> 
            { p with 
                InputFile = docName
                OutputFile = Some <| pdfName 
            })
    with ex -> assertMandatory <| sprintf "COVER SHEET FAILED\n%s" (ex.ToString())


let electricalSchematic (dir:string) : unit = 
    let name1 = System.IO.DirectoryInfo(dir).Name
    Trace.tracefn "Electrical schematic: %s" name1
    match tryFindExactlyOneMatchingFile "*.pdf" dir with
    | Some pdf -> 
        let name1 = System.IO.FileInfo(pdf).Name
        let destFilePath = _outputRoot @@ sprintf "ELECD %s" name1
        Trace.tracefn "destPath: %s" destFilePath
        optionalCopyFile destFilePath pdf
    | None -> assertOptional (sprintf "No diagram for '%s'" dir)


let makeElectricalSchematics () : unit = 
    let dir = _filestoreRoot @@ @"1. Electrical Schematics T7350"
    System.IO.Directory.GetDirectories(dir) |> Array.iter electricalSchematic


let ioSchedule (filePath:string) : unit = 
    let name1 = System.IO.FileInfo(filePath).Name
    let scheduleName = 
        between name1 "59627 - YW Filey STW PLC Asset Replacement - " " - As Built.xlsx"
    Trace.tracefn "IO Schedule: '%s'" scheduleName
    if System.IO.File.Exists(filePath) then 
        let outputFile = 
            _outputRoot @@ sprintf "IOS %s.pdf" scheduleName
        XlsToPdf (fun p -> 
                    { p with 
                        InputFile = filePath
                        OutputFile = Some <| outputFile
                    })
    else 
        assertOptional (sprintf "No io-schedule for '%s'" filePath)

let makeIoSchedules () : unit = 
    let dir = _filestoreRoot @@ @"2. I-O Schedules"
    System.IO.Directory.GetFiles(dir) |> Array.iter ioSchedule


let fdSpec (filePath:string) : unit = 
    let name1 = System.IO.FileInfo(filePath).Name
    let fdsName = 
        between name1 "59627 - YW Filey STW PLC Asset Replacement "  " - As Built.doc"
    Trace.tracefn "FDS: '%s'" fdsName
    if System.IO.File.Exists(filePath) then 
        let outputFile = 
            _outputRoot @@ sprintf "FDS %s.pdf" fdsName
        DocToPdf (fun p -> 
                    { p with 
                        InputFile = filePath
                        OutputFile = Some <| outputFile
                    })
    else 
        assertOptional (sprintf "No FDS for '%s'" filePath)

let makeFDSpecs () : unit = 
    let dir = _filestoreRoot @@ @"3. Functional Design Specifications"
    System.IO.Directory.GetFiles(dir) |> Array.iter fdSpec

let makeTestSpecs () : unit = 
    let fun1 (src:string) : unit = 
        Trace.tracefn "Test Spec: '%s'" src
        optionalCopyFile _outputRoot src
    let dir = _filestoreRoot @@ @"4. Test Specifications"
    match findAllMatchingFiles "*.pdf" dir with
    | [] -> assertOptional (sprintf "No Test Specs for '%s'" dir)  
    | xs -> List.iter fun1 xs


let makeComms () : unit = 
    let pdfMove (src:string) : unit = 
        let name1 = System.IO.FileInfo(src).Name
        let destFilePath = _outputRoot @@ sprintf "COMMS %s" name1
        Trace.tracefn "destPath: %s" destFilePath
        optionalCopyFile destFilePath src
    let xlsPdf (src:string) : unit = 
        let name1 = System.IO.FileInfo(src).Name
        let docName = 
            between name1 "59627 " " - As Built."
        let outputFile = 
            _outputRoot @@ sprintf "COMMS %s.pdf" docName
        XlsToPdf (fun p -> 
                    { p with 
                        InputFile = src
                        OutputFile = Some <| outputFile
                    })

    let dir = _filestoreRoot @@ @"5. Network & Communications"
    match findAllMatchingFiles "*.pdf" dir with
    | [] -> assertOptional (sprintf "No pdfs for '%s'" dir)  
    | xs -> List.iter pdfMove xs    
    match findAllMatchingFiles "*.xls*" dir with
    | [] -> assertOptional (sprintf "No xls for '%s'" dir)  
    | xs -> List.iter xlsPdf xs  



let finalGlobs : string list = 
    [ "*cover-sheet.pdf" 
    ; "ELECD*.pdf" 
    ; "IOS*.pdf" 
    ; "FDS*.pdf"
    ; "SAT*.pdf" 
    ; "COMMS*.pdf" 
    ]

let makeFinal () : unit = 
    let files:string list= 
        List.collect (fun glob -> findAllMatchingFiles glob assetOutputDir) finalGlobs
    PdfConcat (fun p -> 
        { p with 
            OutputFile = makeAssetOutputName "%s S3821 PLC Asset Replacement.pdf"
            PrintQuality = PdfWhatever 
            })
        files






let main () : unit = 
    makeClean () 
    makeOutputDirectory ()
    makeCover () 
    makeElectricalSchematics () 
    makeIoSchedules () 
    makeFDSpecs ()
    makeTestSpecs () 
    makeComms () 
    makeFinal () 


let test01 () = 
    printfn "%b" <| isPrefixOf "Big Bad Wolf" "Bad"
    printfn "%b" <| isPrefixOf "Big Bad Wolf" "Big"
    printfn "%b" <| isSuffixOf "Big Bad Wolf" "Wol"
    printfn "%b" <| isSuffixOf "Big Bad Wolf" "olf"

// Trim() is non-destructive
let dummy02 () = 
    let s1 = " Sample    "
    printfn "s1:'%s'" s1
    let s2 = s1.Trim()
    printfn "s1:'%s'; s2:'%s'" s1 s2

let test02 () = 
    let s1 = "Inside the ships" 
    printfn "leftOf - s1:'%s'; s2:'%s'"  s1 <| leftOf "Inside the ships" " ships"
    printfn "rightOf - s1:'%s'; s2:'%s'" s1 <| rightOf "Inside the ships" "Inside "