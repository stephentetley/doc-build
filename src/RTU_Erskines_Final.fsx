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


#I @"..\packages\Magick.NET-Q8-AnyCPU.7.4.6\lib\net40"
#r @"Magick.NET-Q8-AnyCPU.dll"
open ImageMagick




#load @"DocMake\Base\Common.fs"
#load @"DocMake\Base\OfficeUtils.fs"
#load @"DocMake\Base\FakeLike.fs"
#load @"DocMake\Base\ImageMagickUtils.fs"
#load @"DocMake\Base\SimpleDocOutput.fs"
open DocMake.Base.Common
open DocMake.Base.FakeLike

#load @"DocMake\Builder\BuildMonad.fs"
#load @"DocMake\Builder\Basis.fs"
#load @"DocMake\Builder\WordHooks.fs"
#load @"DocMake\Builder\ExcelHooks.fs"
#load @"DocMake\Builder\PowerPointHooks.fs"
#load @"DocMake\Builder\GhostscriptHooks.fs"
#load @"DocMake\Builder\PdftkHooks.fs"
open DocMake.Base.Common
open DocMake.Base.FakeLike
open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis



#load @"DocMake\Tasks\DocFindReplace.fs"
#load @"DocMake\Tasks\XlsFindReplace.fs"
#load @"DocMake\Tasks\DocToPdf.fs"
#load @"DocMake\Tasks\XlsToPdf.fs"
#load @"DocMake\Tasks\PptToPdf.fs"
#load @"DocMake\Tasks\PdfConcat.fs"
#load @"DocMake\Tasks\PdfRotate.fs"
#load @"DocMake\Tasks\DocPhotos.fs"
#load @"DocMake\FullBuilder.fs"
open DocMake.FullBuilder
open DocMake.Tasks



let _inputRoot = @"G:\work\Projects\rtu\final-docs\input\Erskines_02"
let _outputDir = @"G:\work\Projects\rtu\final-docs\output\erskine-output\batch02"

let generate1 (dir:string) : FullBuild<unit> = 
    match tryFindExactlyOneMatchingFile "*Erskine*.doc*" dir with
    | None -> printfn "BAD:  %s" dir; breturn ()
    | Some ans ->
        buildMonad { 
            let name1 = System.IO.FileInfo(ans).Directory.Name
            printfn "Processing %s ..." name1
            let pdfName = sprintf "%s Erskine Battery Asset Replacement.pdf" name1
            printfn "Output file: %s" pdfName
            let! doc1 = getDocument ans
            let! final = docToPdf doc1 >>= renameTo pdfName
            return ()
        }

let buildScript () : FullBuild<unit>  = 
    let childDirs = System.IO.Directory.GetDirectories(_inputRoot) |> Array.toList
    mapMz generate1 childDirs


let main () : unit = 
    let gsExe = @"C:\programs\gs\gs9.15\bin\gswin64c.exe"
    let pdftkExe = @"C:\programs\PDFtk Server\bin\pdftk.exe"
    let hooks = fullBuilderHooks gsExe pdftkExe

    let env = 
        { WorkingDirectory = _outputDir
          PrintQuality = DocMakePrintQuality.PqScreen
          PdfQuality = PdfPrintSetting.PdfScreen }

    consoleRun env hooks (buildScript ()) 


