// Office deps
#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.Word"
#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Excel\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.Excel"
#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.PowerPoint\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.PowerPoint"
#I @"C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c"
#r "office"

open System.IO


#I @"..\packages\Magick.NET-Q8-AnyCPU.7.3.0\lib\net40"
#r @"Magick.NET-Q8-AnyCPU.dll"
open ImageMagick


// FAKE is local to the project file
#I @"..\packages\FAKE.5.0.0-beta005\tools"
#r @"..\packages\FAKE.5.0.0-beta005\tools\FakeLib.dll"
open Fake
open Fake.Core
open Fake.Core.Environment
open Fake.Core.Globbing.Operators
open Fake.Core.TargetOperators


#load @"DocMake\Base\Common.fs"
#load @"DocMake\Base\Fake.fs"
#load @"DocMake\Base\Office.fs"
#load @"DocMake\Base\CopyRename.fs"
#load @"DocMake\Base\ImageMagick.fs"
open DocMake.Base.Common
open DocMake.Base.Fake
open DocMake.Base.CopyRename
open DocMake.Base.ImageMagick


#load @"DocMake\Tasks\DocToPdf.fs"
#load @"DocMake\Tasks\DocPhotos.fs"
#load @"DocMake\Tasks\PdfConcat.fs"
open DocMake.Tasks.DocToPdf
open DocMake.Tasks.DocPhotos
open DocMake.Tasks.PdfConcat

let _filestoreRoot  = @"G:\work\Projects\barriers\final-docs\Input_Batch01"
let _outputRoot     = @"G:\work\Projects\barriers\final-docs\output"

let siteName = environVarOrDefault "sitename" @"MISSING"


let cleanName           = safeName siteName
let siteInputDir        = _filestoreRoot @@ cleanName
let siteOutputDir       = _outputRoot @@ cleanName


let makeSiteOutputName (fmt:Printf.StringFormat<string->string>) : string = 
    siteOutputDir @@ sprintf fmt cleanName

Target.Create "Clean" (fun _ -> 
    if Directory.Exists(siteOutputDir) then 
        Trace.tracefn " --- Clean folder: '%s' ---" siteOutputDir
        Fake.IO.Directory.delete siteOutputDir
    else 
        Trace.tracefn " --- Clean --- : folder does not exist '%s' ---" siteOutputDir
)


Target.Create "OutputDirectory" (fun _ -> 
    Trace.tracefn " --- Output folder: '%s' ---" siteOutputDir
    maybeCreateDirectory siteOutputDir 
)

Target.Create "SiteWorks" (fun _ -> 
    match tryFindExactlyOneMatchingFile "*Site*Works*.doc*" siteInputDir with
    | Some source -> 
        let pdfOutput = makeSiteOutputName "%s site-works.pdf"
        DocToPdf (fun p -> 
            { p with 
                InputFile = source
                OutputFile = Some <| pdfOutput
            })
    | None -> assertMandatory "SiteWorks: No match"
)

Target.Create "Photos" (fun _ -> 
    let photoSrcPath    = siteInputDir @@ @"PHOTOS"
    let photoDestPath   = siteOutputDir @@ @"Photos"
    maybeCreateDirectory photoDestPath
    multiCopyGlob (photoSrcPath, "*.jpg") photoDestPath
    optimizePhotos photoDestPath

    let docName = makeSiteOutputName "%s extra-photos.docx" 
    let pdfName = pathChangeExtension docName "pdf"
    DocPhotos (fun p -> 
        { p with 
            InputPaths = [ photoDestPath ]            
            OutputFile = docName
            ShowFileName = true 
        })

    DocToPdf (fun p -> 
        { p with 
            InputFile = docName
            OutputFile = Some <| pdfName 
        })

)


let finalGlobs : string list = 
    [ "*site-works.pdf"
    ; "*extra-photos.pdf" ]


Target.Create "Final" (fun _ -> 
    let files:string list= 
        List.collect (fun glob -> findAllMatchingFiles glob siteOutputDir) finalGlobs
    PdfConcat (fun p -> 
        { p with 
            OutputFile = makeSiteOutputName "%s S3953 IS Barrier Replacement.pdf" })
        files
)

// *** Dependencies ***
"Clean"
    ==> "OutputDirectory"
    ==> "SiteWorks"
    ==> "Photos"
    ==> "Final"

Target.RunOrDefault "Final"