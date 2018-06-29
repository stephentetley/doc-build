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


open System.IO



#load @"DocMake\Base\Common.fs"
#load @"DocMake\Base\FakeLike.fs"
#load @"DocMake\Base\ImageMagickUtils.fs"
#load @"DocMake\Base\OfficeUtils.fs"
#load @"DocMake\Base\SimpleDocOutput.fs"
#load @"DocMake\Builder\BuildMonad.fs"
#load @"DocMake\Builder\Document.fs"
#load @"DocMake\Builder\Basis.fs"
#load @"DocMake\Builder\GhostscriptHooks.fs"
#load @"DocMake\Builder\PdftkHooks.fs"
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
#load @"DocMake\Tasks\PdfConcat.fs"
#load @"DocMake\Tasks\PdfRotate.fs"
#load @"DocMake\Tasks\DocPhotos.fs"
#load @"DocMake\FullBuilder.fs"
open DocMake.FullBuilder
open DocMake.Tasks


let _filestoreRoot  = @"G:\work\Projects\flow2\final-docs\Input\Batch02"
let _outputRoot     = @"G:\work\Projects\flow2\final-docs\output"
let _templateRoot   = @"G:\work\Projects\flow2\final-docs\__Templates"


let siteName = "CHESTNUT AVENUE/SPS"
let matches1 : SearchList = 
    [ "#SITENAME", "CHESTNUT AVENUE/SPS"
    ; "#SAINUM", "SAI00004009"
    ]

let cleanName           = safeName siteName
let siteInputDir        = _filestoreRoot </> cleanName
let siteOutputDir       = _outputRoot </> cleanName


let makeSiteOutputName (fmt:Printf.StringFormat<string->string>) : string = 
    siteOutputDir </> sprintf fmt cleanName




// This should be a mandatory task
let cover (matches:SearchList) : FullBuild<PdfDoc> = 
    buildMonad { 
        let templatePath = _templateRoot </> "FC2 Cover TEMPLATE.docx"
        let! template = getTemplateDoc templatePath
        let! d1 = docFindReplace matches template >>= docToPdf >>= renameTo "cover-sheet.pdf"
        return d1 }



let photosDoc (docTitle:string) (jpegSrcPath:string) (pdfName:string) : FullBuild<PdfDoc> = 
    let photoOpts:DocPhotos.DocPhotosOptions = 
        { DocTitle = Some docTitle; ShowFileName = true; CopyToSubDirectory = "Photos" } 

    buildMonad { 
        let! d1 = docPhotos photoOpts [jpegSrcPath]
        let! d2 = breturn d1 >>= docToPdf >>= renameTo pdfName
        return d2
        }
    

let scopeOfWorks () : BuildMonad<'res,PdfDoc> = 
    match tryFindExactlyOneMatchingFile "*Scope of Works*.pdf*" siteInputDir with
    | Some source -> copyToWorkingDirectory source
    | None -> throwError "NO SCOPE OF WORKS"



let citWork () : FullBuild<PdfDoc list> = 
    let proc (glob:string) (renamer:int->string) : FullBuild<PdfDoc list> = 
        findAllMatchingFiles glob siteInputDir |>
            foriM (fun i source ->  
                        copyToWorkingDirectory source >>= renameTo (renamer (i+1)))

    buildMonad {
        let! ds1 = proc "2018-S4371*.pdf" (sprintf "electricals-%03i.pdf")
        let! ds2 = proc "*Prop Works*.pdf" (sprintf "cad-drawing-%03i.pdf")
        return ds1 @ ds2 
    }



    // If this isn't thunkified it will launch Excel when the code is loaded in FSI
let installSheets () : FullBuild<PdfDoc list> = 
    let pdfGen (glob:string) (warnMsg:string) : FullBuild<PdfDoc list> = 
        match findAllMatchingFiles glob siteInputDir with
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


let buildScript (siteName:string) : FullBuild<unit> = 
    
    buildMonad { 
        do! IOActions.clean () >>. IOActions.createOutputDirectory ()
        let! p1 = cover matches1
        let surveyJpegsPath = siteInputDir </> "Survey_Photos"
        let! p2 = photosDoc "Survey Photos" surveyJpegsPath "survey-photos.pdf"
        let! p3 = makePdf "scope-of-works.pdf"  <| scopeOfWorks () 
        let! ps1 = citWork ()
        let! ps2 = installSheets ()
        let surveyJpegsPath = siteInputDir </> "Install_Photos"
        let! pZ = photosDoc "Install Photos" surveyJpegsPath "install-photos.pdf"
        let (pdfs:PdfDoc list) = p1 :: p2 :: p3 :: (ps1 @ ps2 @ [pZ])
        let! (final:PdfDoc) = pdfConcat pdfs >>= renameTo "FINAL.pdf"
        return ()                 
    }


let main () : unit = 
    let env = 
        { WorkingDirectory = siteOutputDir
          PrintQuality = PrintQuality.PqScreen
          PdfQuality = PdfPrintQuality.PdfScreen }

    let appConfig : FullBuildConfig = 
        { GhostscriptPath = @"C:\programs\gs\gs9.15\bin\gswin64c.exe"
          PdftkPath = @"C:\programs\PDFtk Server\bin\pdftk.exe" } 

    runFullBuild env appConfig (buildScript siteName) 


