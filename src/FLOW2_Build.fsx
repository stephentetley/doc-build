// Office deps
#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.Word"
#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Excel\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.Excel"
#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.PowerPoint\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.PowerPoint"
#I @"C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c"
#r "office"


#I @"..\packages\Magick.NET-Q8-AnyCPU.7.3.0\lib\net40"
#r @"Magick.NET-Q8-AnyCPU.dll"
open ImageMagick

#I @"..\packages\Newtonsoft.Json.10.0.3\lib\net45"
#r "Newtonsoft.Json"
open Newtonsoft.Json

open System.IO

// FAKE is local to the project file
#I @"..\packages\FAKE.5.0.0-beta005\tools"
#r @"..\packages\FAKE.5.0.0-beta005\tools\FakeLib.dll"
open Fake
open Fake.Core
open Fake.Core.Environment
open Fake.Core.Globbing.Operators
open Fake.Core.TargetOperators


#load @"DocMake\Base\Common.fs"
#load @"DocMake\Base\FakeExtras.fs"
#load @"DocMake\Base\ImageMagickUtils.fs"
#load @"DocMake\Base\OfficeUtils.fs"
#load @"DocMake\Base\SimpleDocOutput.fs"
#load @"DocMake\Builder\BuildMonad.fs"
#load @"DocMake\Builder\Basis.fs"
#load @"DocMake\Builder\Builders.fs"
open DocMake.Base.Common
open DocMake.Base.FakeExtras
open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis
open DocMake.Builder.Builders

#load @"DocMake\Lib\DocFindReplace.fs"
#load @"DocMake\Lib\DocPhotos.fs"
#load @"DocMake\Lib\DocToPdf.fs"
#load @"DocMake\Lib\XlsToPdf.fs"
#load @"DocMake\Lib\PdfConcat.fs"
open DocMake.Lib.DocFindReplace
open DocMake.Lib.DocPhotos
open DocMake.Lib.DocToPdf
open DocMake.Lib.XlsToPdf
open DocMake.Lib.PdfConcat

// TODO - localize these

let _filestoreRoot  = @"G:\work\Projects\flow2\final-docs\Input\Batch02"
let _outputRoot     = @"G:\work\Projects\flow2\final-docs\output"
let _templateRoot   = @"G:\work\Projects\flow2\final-docs\__Templates"


let siteName = "CHESTNUT AVENUE/SPS"
let matches1 : SearchList = 
    [ "#SITENAME", "CHESTNUT AVENUE/SPS"
    ; "#SAINUM", "SAI00004009"
    ]

let cleanName           = safeName siteName
let siteInputDir        = _filestoreRoot @@ cleanName
let siteOutputDir       = _outputRoot @@ cleanName


let makeSiteOutputName (fmt:Printf.StringFormat<string->string>) : string = 
    siteOutputDir @@ sprintf fmt cleanName

let clean : BuildMonad<'res, unit> =
    if Directory.Exists(siteOutputDir) then 
        tellLine (sprintf " --- Clean folder: '%s' ---" siteOutputDir) >>.
        executeIO (fun () -> Fake.IO.Directory.delete siteOutputDir)
    else 
        tellLine <| sprintf " --- Clean --- : folder does not exist '%s' ---" siteOutputDir


let outputDirectory : BuildMonad<'res, unit> =
    tellLine (sprintf  " --- Output folder: '%s' ---" siteOutputDir) >>.
    executeIO (fun () -> maybeCreateDirectory siteOutputDir)



// This should be a mandatory task
let cover (matches:SearchList) : BuildMonad<'res, PdfDoc> = 
    execWordBuild ( 
        buildMonad { 
            let templatePath = _templateRoot @@ "FC2 Cover TEMPLATE.docx"
            let! template = getTemplate templatePath
            let! d1 = docFindReplace matches template >>= docToPdf >>= renameTo "cover-sheet.pdf"
            return d1 }) 



let photosDoc (docTitle:string) (jpegSrcPath:string) (pdfName:string) : BuildMonad<'res, PdfDoc> = 
    execWordBuild <| 
        buildMonad { 
            let! d1 = photoDoc (Some docTitle) true [jpegSrcPath]
            let! d2 = breturn d1 >>= docToPdf >>= renameTo pdfName
            return d2
            }
    

let scopeOfWorks () : BuildMonad<'res,PdfDoc> = 
    match tryFindExactlyOneMatchingFile "*Scope of Works*.pdf*" siteInputDir with
    | Some source -> copyToWorkingDirectory source
    | None -> throwError "NO SCOPE OF WORKS"



let citWork () : BuildMonad<'res, PdfDoc list> = 
    let proc (glob:string) (renamer:int->string) : BuildMonad<'res, PdfDoc list> = 
        findAllMatchingFiles glob siteInputDir |>
            foriM (fun i source ->  
                        copyToWorkingDirectory source >>= renameTo (renamer (i+1)))

    buildMonad {
        let! ds1 = proc "2018-S4371*.pdf" (sprintf "electricals-%03i.pdf")
        let! ds2 = proc "*Prop Works*.pdf" (sprintf "cad-drawing-%03i.pdf")
        return ds1 @ ds2 
    }



    // If this isn't thunkified it will launch Excel when the code is loaded in FSI
let installSheets () : BuildMonad<'res, PdfDoc list> = 
    let pdfGen (glob:string) (warnMsg:string) : ExcelBuild<PdfDoc list> = 
        match findAllMatchingFiles glob siteInputDir with
        | [] -> 
            tellLine warnMsg >>. breturn []
             
        | xs -> 
            forM xs (fun path -> 
                printfn "installSheet: %s" path
                getDocument path >>= xlsToPdf true )
    
    withNameGen (sprintf "install-%03i.pdf") << execExcelBuild <| 
        buildMonad { 
            let! ds1 = pdfGen "*Flow meter*.xls*" "NO FLOW METER INSTALL SHEETS"
            let! ds2 = pdfGen "*Pressure inst*.xls*" "NO PRESSURE SENSOR INSTALL SHEET"
            return ds1 @ ds2
            }


// *******************************************************


let buildScript (siteName:string) : BuildMonad<'res,unit> = 
    let gsExe = @"C:\programs\gs\gs9.15\bin\gswin64c.exe"
    buildMonad { 
        do! clean >>. outputDirectory
        let! p1 = cover matches1
        let surveyJpegsPath = siteInputDir @@ "Survey_Photos"
        let! p2 = photosDoc "Survey Photos" surveyJpegsPath "survey-photos.pdf"
        let! p3 = makePdf "scope-of-works.pdf"  <| scopeOfWorks () 
        let! ps1 = citWork ()
        let! ps2 = installSheets ()
        let surveyJpegsPath = siteInputDir @@ "Install_Photos"
        let! pZ = photosDoc "Install Photos" surveyJpegsPath "install-photos.pdf"
        let (pdfs:PdfDoc list) = p1 :: p2 :: p3 :: (ps1 @ ps2 @ [pZ])
        let! (final:PdfDoc) = execGsBuild gsExe (pdfConcat pdfs) >>= renameTo "FINAL.pdf"
        return ()                 
    }


let main () : unit = 
    let env = 
        { WorkingDirectory = siteOutputDir
          PrintQuality = DocMakePrintQuality.PqScreen
          PdfQuality = PdfPrintSetting.PdfScreen }

    consoleRun env (buildScript siteName) 


