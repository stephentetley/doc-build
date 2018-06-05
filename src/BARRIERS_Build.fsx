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

// FAKE dependencies are getting onorous...
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
open Fake
open Fake.IO.FileSystemOperators



#load @"DocMake\Base\Common.fs"
#load @"DocMake\Base\FakeExtras.fs"
#load @"DocMake\Base\ImageMagickUtils.fs"
#load @"DocMake\Base\OfficeUtils.fs"
#load @"DocMake\Base\SimpleDocOutput.fs"
#load @"DocMake\Builder\BuildMonad.fs"
#load @"DocMake\Builder\Basis.fs"
#load @"DocMake\Builder\WordBuilder.fs"
#load @"DocMake\Builder\ExcelBuilder.fs"
#load @"DocMake\Builder\GhostscriptBuilder.fs"
open DocMake.Base.Common
open DocMake.Base.FakeExtras
open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis
open DocMake.Builder.WordBuilder
open DocMake.Builder.GhostscriptBuilder

#load @"DocMake\Lib\DocPhotos.fs"
#load @"DocMake\Lib\DocToPdf.fs"
#load @"DocMake\Lib\PdfConcat.fs"
open DocMake.Lib

// Output is just "Site Works" doc and the collected "Photo doc"


let clean () : BuildMonad<'res, unit> =
    buildMonad { 
        let! cwd = askWorkingDirectory ()
        if Directory.Exists(cwd) then 
            do printfn " --- Clean folder: '%s' ---" cwd
            do! deleteWorkingDirectory ()
        else 
            do printfn " --- Clean --- : folder does not exist '%s' ---" cwd
    }



let outputDirectory () : BuildMonad<'res, unit> =
    buildMonad { 
        let! cwd = asksEnv (fun e -> e.WorkingDirectory)
        do printfn  " --- Output folder: '%s' ---" cwd
        do! createWorkingDirectory ()
    }


// No cover needed

let siteWorks (siteInputDir:string) : BuildMonad<'res,PdfDoc> = 
    match tryFindExactlyOneMatchingFile "*Site Works*.doc*" siteInputDir with
    | Some source -> 
            execWordBuild <| (getDocument source >>= docToPdf)
    | None -> throwError "No Site Works"


let photosDoc (docTitle:string) (jpegSrcPath:string) : BuildMonad<'res, PdfDoc> = 
    execWordBuild <| 
        (photoDoc (Some docTitle) true [jpegSrcPath] >>= docToPdf)
    




// *******************************************************


let buildScript (inputRoot:string) (siteName:string) : BuildMonad<'res,PdfDoc> = 
    let gsExe = @"C:\programs\gs\gs9.15\bin\gswin64c.exe"
    let cleanName           = safeName siteName
    let siteInputDir        = inputRoot @@ cleanName
    let jpegsSrcPath        = siteInputDir @@ "PHOTOS"
    let finalName           = sprintf "%s S3953 IS Barrier Replacement.pdf" cleanName
    localSubDirectory cleanName <| 
        buildMonad { 
            do! clean () >>. outputDirectory ()
            let! p1 = makePdf "site-works.pdf"          <| siteWorks siteInputDir
            let! p2 = makePdf "site-work-photos.pdf"    <| photosDoc "Site Work Photos" jpegsSrcPath 
            let pdfs = [p1;p2]
            let! (final:PdfDoc) = makePdf finalName     <| execGsBuild gsExe (pdfConcat pdfs)
            return final            
        }


let getSites (root:string) : string [] = 
    let slashName (name:string) = name.Replace(oldChar='_', newChar='/')
    let getName (path:string) = 
        slashName <| System.IO.DirectoryInfo(path).Name
    System.IO.Directory.GetDirectories(root) 
        |> Array.map getName


let main () : unit = 
    let inputRoot      = @"G:\work\Projects\barriers\final-docs\input\Batch02"
    let outputRoot     = @"G:\work\Projects\barriers\final-docs\output\Batch02"
    let env = 
        { WorkingDirectory = outputRoot
          PrintQuality = DocMakePrintQuality.PqScreen
          PdfQuality = PdfPrintSetting.PdfScreen }
    let siteList = getSites inputRoot |> Array.toList 
    consoleRun env <| forMz siteList (buildScript inputRoot) 



