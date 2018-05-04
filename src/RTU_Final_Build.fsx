// Run in PowerShell not fsi:
// PS> cd <path-to-src>
// PS> ..\packages\FAKE.5.0.0-beta005\tools\FAKE.exe .\RTU_Final_Build.fsx Dummy

// With params:
// PS> ..\packages\FAKE.5.0.0-beta005\tools\FAKE.exe .\RTU_Final_Build.fsx Final --envar sitename="HELLO/WORLD"

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
#load @"DocMake\Base\JsonUtils.fs"
#load @"DocMake\Base\OfficeUtils.fs"
#load @"DocMake\Base\SimpleDocOutput.fs"
open DocMake.Base.Common
open DocMake.Base.FakeExtras
open DocMake.Base.JsonUtils

#load @"DocMake\Tasks\DocFindReplace.fs"
#load @"DocMake\Tasks\DocPhotos.fs"
#load @"DocMake\Tasks\DocToPdf.fs"
#load @"DocMake\Tasks\PdfConcat.fs"
open DocMake.Tasks.DocFindReplace
open DocMake.Tasks.DocPhotos
open DocMake.Tasks.DocToPdf
open DocMake.Tasks.PdfConcat

// NOTE - can generate a batch file to do "many-to-one"
// We usually have many final docs to make (many sites) but the style of 
// Fake is to make one agglomerate out of many parts.
// Generating a batch file that invokes Fake for each site solves this.

let _filestoreRoot  = @"G:\work\Projects\rtu\Final_Docs\batch2_input"
let _outputRoot     = @"G:\work\Projects\rtu\Final_Docs\batch2_output"
let _templateRoot   = @"G:\work\Projects\rtu\Final_Docs\__Templates"
let _jsonRoot       = @"G:\work\Projects\rtu\Final_Docs\__Json"

// siteName is an envVar so we can use this build script to build many 
// sites (they all follow the same directory/file structure).
let siteName = environVarOrDefault "sitename" @"MISSING"


let cleanName           = safeName siteName
let siteInputDir        = _filestoreRoot @@ cleanName
let siteOutputDir       = _outputRoot @@ cleanName


let makeSiteOutputName (fmt:Printf.StringFormat<string->string>) : string = 
    siteOutputDir @@ sprintf fmt cleanName

Target.Create "Clean" (fun _ -> 
    if Directory.Exists siteOutputDir then 
        Trace.tracefn " --- Clean folder: '%s' ---" siteOutputDir
        Fake.IO.Directory.delete siteOutputDir
    else 
        Trace.tracefn " --- Clean --- : folder does not exist '%s' ---" siteOutputDir
)


Target.Create "OutputDirectory" (fun _ -> 
    Trace.tracefn " --- Output folder: '%s' ---" siteOutputDir
    maybeCreateDirectory siteOutputDir
)

Target.Create "CoverSheet" (fun _ ->
    let template = _templateRoot @@ "MM3x-to-MMIM RTU Cover Sheet.docx"
    let jsonSource = _jsonRoot @@ (sprintf "%s_findreplace.json" cleanName)
    let docName = makeSiteOutputName "%s Cover Sheet.docx"
    let pdfName = pathChangeExtension docName "pdf"

    Trace.tracefn " --- Cover sheet for: %s --- " siteName
    if File.Exists(jsonSource) then
        let matches = readJsonStringPairs jsonSource
        DocFindReplace (fun p -> 
        
            { p with 
                TemplateFile = template
                OutputFile = docName
                Matches  = matches 
            }) 
    
        DocToPdf (fun p -> 
            { p with 
                InputFile = docName
                OutputFile = Some <| pdfName 
            })
    else assertMandatory "CoverSheet failed no json matches"
)



Target.Create "SurveySheet" (fun _ ->
    match tryFindExactlyOneMatchingFile "*urvey.doc*" siteInputDir with
    | Some docFile -> 
        let pdfFile = makeSiteOutputName "%s Survey Sheet.pdf" 
        DocToPdf (fun p -> 
        { p with 
            InputFile = docFile
            OutputFile = Some <| pdfFile
        })
    | None -> assertOptional "NO SURVEY SHEET"
)


Target.Create "SurveyPhotos" (fun _ ->
    let photosPath = siteInputDir @@ "Survey Photos"
    let docName = makeSiteOutputName "%s Survey Photos.docx" 
    let pdfName = pathChangeExtension docName "pdf"

    if System.IO.Directory.Exists(photosPath) then
        DocPhotos (fun p -> 
            { p with 
                InputPaths = [photosPath]            
                OutputFile = docName
                ShowFileName = true 
            })
        DocToPdf (fun p -> 
            { p with 
                InputFile = docName
                OutputFile = Some <| pdfName 
            })
    else assertOptional "NO SURVEY PHOTOS"
)

Target.Create "InstallSheet" (fun _ ->
    match tryFindExactlyOneMatchingFile "*Site Works*.doc*" siteInputDir with
    | Some docName -> 
        let pdfName = makeSiteOutputName "%s Install Sheet.pdf" 
        DocToPdf (fun p -> 
            { p with 
                InputFile = docName
                OutputFile = Some <| pdfName
            })
    | None -> assertMandatory "NO INSTALL SHEET"
)

Target.Create "InstallPhotos" (fun _ ->
    let photosPath = siteInputDir @@ "install photos"
    let docName = makeSiteOutputName "%s Install Photos.docx" 
    let pdfName = pathChangeExtension docName "pdf"

    if System.IO.Directory.Exists(photosPath) then
        DocPhotos (fun p -> 
            { p with 
                InputPaths = [photosPath]            
                OutputFile = docName
                ShowFileName = true 
            })
        DocToPdf (fun p -> 
            { p with 
                InputFile = docName
                OutputFile = Some <| pdfName 
            })
    else assertOptional "NO INSTALL PHOTOS"
)

let finalGlobs : string list = 
    [ "* Cover Sheet.pdf" ;
      "* Survey Sheet.pdf" ;
      "* Survey Photos.pdf" ;
      "* Install Sheet.pdf" ;
      "* Install Photos.pdf" ]

Target.Create "Final" (fun _ ->
    let files:string list= 
        List.collect (fun glob -> findAllMatchingFiles glob siteOutputDir) finalGlobs
    PdfConcat (fun p -> 
        { p with 
            OutputFile = makeSiteOutputName "%s S3953 RTU Asset Replacement.pdf" })
        files
)
// *** Dummy cases

Target.Create "Dummy" (fun _ ->
    printfn "Message from Dummy target"
)

Target.Create "None" (fun _ ->
    printfn "None"
)


// *** Dependencies ***
"Clean"
    ==> "OutputDirectory"

"OutputDirectory"
    ==> "CoverSheet"
    ==> "SurveySheet"
    ==> "SurveyPhotos"
    ==> "InstallSheet"
    ==> "InstallPhotos"
    ==> "Final"

// Note seemingly Fake files must end with this...
Target.RunOrDefault "None"
