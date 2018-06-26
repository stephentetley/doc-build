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

// ImageMagick for DocPhotos
#I @"..\packages\Magick.NET-Q8-AnyCPU.7.4.6\lib\net40"
#r @"Magick.NET-Q8-AnyCPU.dll"
open ImageMagick


#I @"..\packages\ExcelProvider.0.8.2\lib"
#r "ExcelProvider.dll"
open FSharp.ExcelProvider

open System.IO

#load @"DocMake\Base\Common.fs"
#load @"DocMake\Base\FakeLike.fs"
#load @"DocMake\Base\ImageMagickUtils.fs"
#load @"DocMake\Base\OfficeUtils.fs"
#load @"DocMake\Base\SimpleDocOutput.fs"
#load @"DocMake\Builder\BuildMonad.fs"
#load @"DocMake\Builder\Basis.fs"
#load @"DocMake\Builder\WordBuilder.fs"
#load @"DocMake\Builder\ExcelBuilder.fs"
#load @"DocMake\Builder\PowerPointBuilder.fs"
#load @"DocMake\Builder\GhostscriptBuilder.fs"
#load @"DocMake\Builder\PdftkBuilder.fs"
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


#load "Proprietry.fs"
open Proprietry

let _templateRoot       = @"G:\work\Projects\usar\final-docs\__Templates"
let _inputRoot          = @"G:\work\Projects\usar\final-docs\input\June2018_INPUT"
let _outputRoot         = @"G:\work\Projects\usar\final-docs\output\June2018"

let makeCoverMatches (siteName:string) (saiLookups:SaiLookups) : option<SearchList> =  
    match getSaiNumber siteName saiLookups with
    | None -> None
    | Some sai -> 
        Some <| 
            [ "#SITENAME",      siteName   
            ; "#SAINUM" ,       sai
            ]



let cover (siteName:string) : FullBuild<PdfDoc> = 
    buildMonad { 
        let templatePath = _templateRoot </> @"USAR Cover Sheet.docx"
        let! template = getTemplateDoc templatePath
        let docOutName = sprintf "%s cover-sheet.docx" (safeName siteName)
        let lookups = getSaiLookups ()
        match makeCoverMatches siteName lookups with
        | Some matches -> 
            let! d1 = docFindReplace matches template >>= renameTo docOutName 
            let! d2 = docToPdf d1
            return d2
        | None -> throwError "cover - no sai number" |> ignore
    }

/// Survey / Disposed of / Install
/// "*Survey*.docx", survey might not be the last word of the name


let processSheets (globPattern:string) (inputPath:string) : FullBuild<PdfDoc list> =
    let surveys = findAllMatchingFiles globPattern inputPath
    mapM (fun survey1 -> getDocument survey1 >>= docToPdf) surveys


let surveySheets (inputPath:string) : FullBuild<PdfDoc list> =
    let glob = "*Survey*.docx"
    processSheets glob inputPath

    
let disposedOfSheets (inputPath:string) : FullBuild<PdfDoc list> =
    let glob = "*Disposed*.docx"
    processSheets glob inputPath

let installSheets (inputPath:string) : FullBuild<PdfDoc list> =
    let glob = "*Install*.docx"
    processSheets glob inputPath


    


let buildScript1 (inputPath:string) : FullBuild<PdfDoc> = 
    let siteName    = FileInfo(inputPath).Name |> (fun s -> s.Replace("_", "/"))
    let cleanName   = safeName siteName
    let finalName   = sprintf "%s S3820 Ultrasonic Asset Replacement.pdf" cleanName
    localSubDirectory cleanName <| 
        buildMonad { 
            // do! clean () >>. outputDirectory ()
            let! p1 = cover siteName
            let! ps2 = surveySheets inputPath
            let! ps3 = disposedOfSheets inputPath
            let! ps4 = installSheets inputPath
            let pdfs : PdfDoc list = p1 :: (ps2 @ ps3 @ ps4)
            let! (final:PdfDoc) = makePdf finalName     <| pdfConcat pdfs
            return final            
        }

let buildScript () : FullBuild<unit> = 
    let inputs = 
        System.IO.Directory.GetDirectories(_inputRoot) 
            |> Array.toList
            
    mapMz buildScript1 inputs

    

let main () : unit = 
    let gsExe = @"C:\programs\gs\gs9.15\bin\gswin64c.exe"
    let pdftkExe = @"C:\programs\PDFtk Server\bin\pdftk.exe"
    let hooks = fullBuilderHooks gsExe pdftkExe    
    let env = 
        { WorkingDirectory = _outputRoot
          PrintQuality = DocMakePrintQuality.PqScreen
          PdfQuality = PdfPrintSetting.PdfScreen }


    consoleRun env hooks (buildScript ()) 
