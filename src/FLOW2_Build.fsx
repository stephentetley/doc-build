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


#load @"DocMake\Base\Common.fs"
#load @"DocMake\Base\FakeFake.fs"
#load @"DocMake\Base\FakeExtras.fs"
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
open DocMake.Base.FakeExtras
open DocMake.Base.FakeFake
open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis


#load @"DocMake\Lib\DocFindReplace.fs"
#load @"DocMake\Lib\DocPhotos.fs"
#load @"DocMake\Lib\DocToPdf.fs"
#load @"DocMake\Lib\XlsToPdf.fs"
#load @"DocMake\Lib\PptToPdf.fs"
#load @"DocMake\Lib\PdfConcat.fs"
#load @"DocMake\FullBuilder.fs"
open DocMake.FullBuilder
open DocMake.Lib


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

let clean : FullBuild<unit> =
    buildMonad { 
        if Directory.Exists(siteOutputDir) then 
            do printfn " --- Clean folder: '%s' ---" siteOutputDir
            do! executeIO (fun () -> Fake.IO.Directory.delete siteOutputDir)
        else 
            do printfn " --- Clean --- : folder does not exist '%s' ---" siteOutputDir
    }

let outputDirectory : FullBuild<unit> =
    buildMonad { 
        do printfn  " --- Output folder: '%s' ---" siteOutputDir
        do! executeIO (fun () -> maybeCreateDirectory siteOutputDir)
    }


// This should be a mandatory task
let cover (matches:SearchList) : FullBuild<PdfDoc> = 
    buildMonad { 
        let templatePath = _templateRoot </> "FC2 Cover TEMPLATE.docx"
        let! template = getTemplate templatePath
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
        do! clean >>. outputDirectory
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
    let gsExe = @"C:\programs\gs\gs9.15\bin\gswin64c.exe"
    let pdftkExe = @"C:\programs\PDFtk Server\bin\pdftk.exe"
    let hooks = fullBuilderHooks gsExe pdftkExe

    let env = 
        { WorkingDirectory = siteOutputDir
          PrintQuality = DocMakePrintQuality.PqScreen
          PdfQuality = PdfPrintSetting.PdfScreen }

    consoleRun env hooks (buildScript siteName) 


