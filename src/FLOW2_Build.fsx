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

// Use FSharp.Data for CSV output (Proprietry.fs)
#I @"..\packages\FSharp.Data.3.0.0-beta3\lib\net45"
#r @"FSharp.Data.dll"

// Use ExcelProvider to read SAI numbers spreadsheet (Proprietry.fs)
#I @"..\packages\ExcelProvider.0.8.2\lib"
#r "ExcelProvider.dll"

open System.IO



#load @"DocMake\Base\Common.fs"
#load @"DocMake\Base\FakeLike.fs"
#load @"DocMake\Base\ImageMagickUtils.fs"
#load @"DocMake\Base\OfficeUtils.fs"
#load @"DocMake\Base\SimpleDocOutput.fs"
#load @"DocMake\Builder\BuildMonad.fs"
#load @"DocMake\Builder\Document.fs"
#load @"DocMake\Builder\Basis.fs"
#load @"DocMake\Builder\ShellHooks.fs"
open DocMake.Base.Common
open DocMake.Base.FakeLike
open DocMake.Builder.BuildMonad
open DocMake.Builder.Document
open DocMake.Builder.Basis

#load @"DocMake\Tasks\IOActions.fs"
#load @"DocMake\Tasks\DocFindReplace.fs"
#load @"DocMake\Tasks\XlsFindReplace.fs"
#load @"DocMake\Tasks\DocToPdf.fs"
#load @"DocMake\Tasks\XlsToPdf.fs"
#load @"DocMake\Tasks\PptToPdf.fs"
#load @"DocMake\Tasks\MdToDoc.fs"
#load @"DocMake\Tasks\PdfConcat.fs"
#load @"DocMake\Tasks\PdfRotate.fs"
#load @"DocMake\Tasks\DocPhotos.fs"
#load @"DocMake\FullBuilder.fs"
open DocMake.FullBuilder
open DocMake.Tasks

#load "Proprietry.fs"
open Proprietry


let _inputRoot      = @"G:\work\Projects\flow2\final-docs\Input\Batch02"
let _outputRoot     = @"G:\work\Projects\flow2\final-docs\Output\Batch02"
let _templateRoot   = @"G:\work\Projects\flow2\final-docs\__Templates"


let makeCoverMatches (siteName:string) (saiLookups:SaiLookups) : option<SearchList> =  
    match getSaiNumber siteName saiLookups with
    | None -> None
    | Some sai -> 
        Some <| 
            [ "#SITENAME",      siteName   
            ; "#SAINUM" ,       sai
            ]







// This should be a mandatory task
let cover (siteName:string) : FullBuild<PdfDoc> = 
    buildMonad { 
        let templatePath = _templateRoot </> "TEMPLATE Flow Confirmation Phase2 Cover.docx"
        let! template = getTemplateDoc templatePath
        let docOutName = sprintf "%s cover-sheet.docx" (underscoreName siteName)
        let lookups = getSaiLookups ()
        match makeCoverMatches siteName lookups with
        | Some matches -> 
            let! d1 = docFindReplace matches template >>= renameTo docOutName 
            let! d2 = docToPdf d1
            return d2
        | None -> throwError "cover - no sai number" |> ignore
    }


// Common procedure for both survey and Install photos
let photosDoc (docTitle:string) (jpegSrcPath:string) (pdfName:string) : FullBuild<PdfDoc> = 
    let photoOpts:DocPhotos.DocPhotosOptions = 
        { DocTitle = Some docTitle; ShowFileName = true; CopyToSubDirectory = "Photos" } 

    buildMonad { 
        let! d1 = docPhotos photoOpts [jpegSrcPath]
        let! d2 = breturn d1 >>= docToPdf >>= renameTo pdfName
        return d2
        }
    

let scopeOfWorks (inputPath:string) : BuildMonad<'res,PdfDoc> = 
    match tryFindExactlyOneMatchingFile "*Scope of Works*.pdf*" inputPath with
    | Some source -> copyToWorkingDirectory source
    | None -> throwError "NO SCOPE OF WORKS"



let citWork (inputPath) : FullBuild<PdfDoc list> = 
    let proc (glob:string) (renamer:int->string) : FullBuild<PdfDoc list> = 
        findAllMatchingFiles glob inputPath |>
            foriM (fun i source ->  
                        copyToWorkingDirectory source >>= renameTo (renamer (i+1)))

    buildMonad {
        let! ds1 = proc "2018-S4371*.pdf" (sprintf "electricals-%03i.pdf")
        let! ds2 = proc "*Prop Works*.pdf" (sprintf "cad-drawing-%03i.pdf")
        return ds1 @ ds2 
    }



    // If this isn't thunkified it will launch Excel when the code is loaded in FSI
let installSheets (inputPath:string) : FullBuild<PdfDoc list> = 
    let pdfGen (glob:string) (warnMsg:string) : FullBuild<PdfDoc list> = 
        match findAllMatchingFiles glob inputPath with
        | [] -> 
            printfn "%s" warnMsg; breturn []
             
        | xs -> 
            forM xs (fun path -> 
                printfn "installSheet: %s" path
                getDocument path >>= xlsToPdf true )
    
    withNameGen (sprintf "install-%03i.pdf") <| 
        buildMonad { 
            let! ds1 = pdfGen "*Flow meter*.xls*" "NO FLOW METER INSTALL SHEETS"
            let! ds2 = pdfGen "*Pressure inst*.xls*" "NO PRESSURE SENSOR INSTALL SHEET"
            return ds1 @ ds2
            }


// *******************************************************


let buildScript1 (inputPath:string) : FullBuild<PdfDoc> = 
    let siteName    = slashName <| FileInfo(inputPath).Name
    let cleanName   = safeName siteName
    localSubDirectory cleanName <|
        buildMonad { 
            do! IOActions.clean () >>. IOActions.createOutputDirectory ()
            let siteInputDir = _inputRoot </> underscoreName siteName
            let! p1 = cover siteName
            let surveyJpegsPath = siteInputDir </> "Survey_Photos"
            let! p2 = photosDoc "Survey Photos" surveyJpegsPath "survey-photos.pdf"
            let! p3 = makePdf "scope-of-works.pdf"  <| scopeOfWorks siteInputDir 
            let! ps1 = citWork siteInputDir
            let! ps2 = installSheets siteInputDir
            let surveyJpegsPath = siteInputDir </> "Install_Photos"
            let! pZ = photosDoc "Install Photos" surveyJpegsPath "install-photos.pdf"
            let (pdfs:PdfDoc list) = p1 :: p2 :: p3 :: (ps1 @ ps2 @ [pZ])
            let! (final:PdfDoc) = pdfConcat pdfs >>= renameTo "FINAL.pdf"
            return final                 
        }

let buildScript () : FullBuild<unit>  = 
    let inputs = 
        System.IO.Directory.GetDirectories(_inputRoot) 
            |> Array.toList
    mapMz buildScript1 inputs

let main () : unit = 
    let env = 
        { WorkingDirectory = _outputRoot
          PrintQuality = PrintQuality.PqScreen
          PdfQuality = PdfPrintQuality.PdfScreen }

    let appConfig : FullBuildConfig = 
        { GhostscriptPath = @"C:\programs\gs\gs9.15\bin\gswin64c.exe"
          PdftkPath = @"C:\programs\PDFtk Server\bin\pdftk.exe"
          PandocPath = @"pandoc" } 

    runFullBuild env appConfig <| buildScript ()


