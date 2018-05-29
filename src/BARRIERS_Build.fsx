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

let _filestoreRoot  = @"G:\work\Projects\barriers\final-docs\input\Batch02"
let _outputRoot     = @"G:\work\Projects\barriers\final-docs\output\Batch02"


// Output is just "Site Works" doc and collected "Photo doc"



let siteName = "BARTON LE WILLOW/STW"


//let cleanName           = safeName siteName
//let siteInputDir        = _filestoreRoot @@ cleanName
//let siteOutputDir       = _outputRoot @@ cleanName


//let makeSiteOutputName (fmt:Printf.StringFormat<string->string>) : string = 
//    siteOutputDir @@ sprintf fmt cleanName


// TODO this should clean (delete) the working directory
let clean : BuildMonad<'res, unit> =
    buildMonad { 
        let! cwd = asksEnv (fun e -> e.WorkingDirectory)
        if Directory.Exists(cwd) then 
            do! tellLine (sprintf " --- Clean folder: '%s' ---" cwd)
            do! deleteWorkingDirectory ()
        else 
            do! tellLine <| sprintf " --- Clean --- : folder does not exist '%s' ---" cwd
    }



let outputDirectory : BuildMonad<'res, unit> =
    buildMonad { 
        let! cwd = asksEnv (fun e -> e.WorkingDirectory)
        do! tellLine (sprintf  " --- Output folder: '%s' ---" cwd)
        do! createWorkingDirectory ()
    }


// No cover needed

let siteWorks (siteInputDir:string) : BuildMonad<'res,PdfDoc> = 
    match tryFindExactlyOneMatchingFile "*Site Works*.doc*" siteInputDir with
    | Some source -> 
            execWordBuild <| (getDocument source >>= docToPdf)
    | None -> throwError "No Site Works"


let photosDoc (docTitle:string) (jpegSrcPath:string) (pdfName:string) : BuildMonad<'res, PdfDoc> = 
    execWordBuild <| 
        buildMonad { 
            let! d1 = photoDoc (Some docTitle) true [jpegSrcPath]
            let! d2 = breturn d1 >>= docToPdf >>= renameTo pdfName
            return d2
            }
    




// *******************************************************


let buildScript (siteName:string) : BuildMonad<'res,unit> = 
    let gsExe = @"C:\programs\gs\gs9.15\bin\gswin64c.exe"
    let cleanName           = safeName siteName
    let siteInputDir        = _filestoreRoot @@ cleanName
    let cwd                 = _outputRoot @@ cleanName

    localWorkingDirectory cwd <| 
        buildMonad { 
            do! clean >>. outputDirectory
            let! p1 = siteWorks siteInputDir
            let surveyJpegsPath = siteInputDir @@ "PHOTOS"
            let! p2 = photosDoc "Survey Photos" surveyJpegsPath "survey-photos.pdf"
            let (pdfs:PdfDoc list) = [p1;p2]
            let! (final:PdfDoc) = execGsBuild gsExe (pdfConcat pdfs) >>= renameTo "FINAL.pdf"
            return ()                 
        }



// TODO - this should be a many-build script.

let main () : unit = 
    let env = 
        { WorkingDirectory = _outputRoot
          PrintQuality = DocMakePrintQuality.PqScreen
          PdfQuality = PdfPrintSetting.PdfScreen }

    consoleRun env (buildScript siteName) 




