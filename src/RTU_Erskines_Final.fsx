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

// Use FSharp.Data for CSV output
#I @"..\packages\FSharp.Data.3.0.0-beta3\lib\net45"
#r @"FSharp.Data.dll"
open FSharp.Data


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

#load @"Proprietry.fs"
open Proprietry


let _inputRoot = @"G:\work\Projects\rtu\Erskines\edms-final-docs\input\input-june2018"
let _outputDir = @"G:\work\Projects\rtu\Erskines\edms-final-docs\output\output-june2018"

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


let uploadReceipt (dirList:string list) : FullBuild<unit> = 
    let siteFromPath (path:string) = 
        slashName <| System.IO.DirectoryInfo(path).Name
        
    let config = 
        { MakeTitle = 
            fun name -> sprintf "%s Erskine Battery Asset Replacement" (name.Replace("/", " "))
          MakeDocName =
            fun name -> sprintf "%s Erskine Battery Asset Replacement.pdf" (name.Replace("/", "_"))
          ConstantParams = 
            { ProjectName = "RTU Asset Replacement"
              ProjectCode = "S3953"
              Category = "O & M Manuals"
              ReferenceNumber = "S3953"
              Revision = "1"
              DocumentDate = standardDocumentDate ()
              SheetVolume = "" } 
        }/// Usually blank
    buildMonad { 
        let siteNames = List.map siteFromPath dirList
        do! makeUploadForm siteNames config
    }

let buildScript () : FullBuild<unit>  = 
    buildMonad { 
        let childDirs = System.IO.Directory.GetDirectories(_inputRoot) |> Array.toList
        do! mapMz generate1 childDirs
        do! uploadReceipt childDirs
    }

let main () : unit = 
    let env = 
        { WorkingDirectory = _outputDir
          PrintQuality = PrintQuality.PqScreen
          PdfQuality = PdfPrintQuality.PdfScreen }

    let appConfig : FullBuildConfig = 
            { GhostscriptPath = @"C:\programs\gs\gs9.15\bin\gswin64c.exe"
              PdftkPath = @"C:\programs\PDFtk Server\bin\pdftk.exe" } 

    runFullBuild env appConfig (buildScript ()) 


