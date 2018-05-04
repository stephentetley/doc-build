// Run in PowerShell not fsi:
// PS> cd <path-to-src>
// PS> ..\packages\FAKE.5.0.0-beta005\tools\FAKE.exe .\TESTCASE_DocMake.fsx Dummy

// With params:
// PS> ..\packages\FAKE.5.0.0-beta005\tools\FAKE.exe .\TESTCASE_DocMake.fsx Dummy --envar sitename="HELLO/WORLD"

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

#I @"..\packages\Newtonsoft.Json.10.0.3\lib\net45"
#r "Newtonsoft.Json"
open Newtonsoft.Json

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
#load @"DocMake\Base\CopyRename.fs"
#load @"DocMake\Base\SimpleDocOutput.fs"
open DocMake.Base.Common
open DocMake.Base.FakeExtras
open DocMake.Base.ImageMagickUtils
open DocMake.Base.JsonUtils
open DocMake.Base.CopyRename
open DocMake.Base.SimpleDocOutput

#load @"DocMake\Tasks\DocPhotos.fs"
#load @"DocMake\Tasks\DocFindReplace.fs"
#load @"DocMake\Tasks\DocToPdf.fs"
#load @"DocMake\Tasks\PdfConcat.fs"
#load @"DocMake\Tasks\PptToPdf.fs"
#load @"DocMake\Tasks\XlsToPdf.fs"
open DocMake.Tasks.DocFindReplace
open DocMake.Tasks.DocPhotos
open DocMake.Tasks.DocToPdf
open DocMake.Tasks.PdfConcat
open DocMake.Tasks.PptToPdf
open DocMake.Tasks.XlsToPdf


// The _Alt script expects input files in a different organization:
//  <root>
//      \... CIT
//          \... Site1 
//          \... Site2 
//      \... SITE_WORK
//          \... Site1
//          \... Site2 
//      \... SURVEYS
//          \... Site1
//          \... Site2 


let _filestoreRoot  = @"G:\work\Projects\samps\final-docs\input\May2018_batch01"
let _outputRoot     = @"G:\work\Projects\samps\final-docs\output\May_Batch01"
let _templateRoot   = @"G:\work\Projects\samps\final-docs\__Templates"
let _jsonRoot       = @"G:\work\Projects\samps\final-docs\__Json"

// siteName is an envVar so we can use this build script to build many 
// sites (they all follow the same directory/file structure).
let siteName = environVarOrDefault "sitename" @"MISSING"

let cleanName           = safeName siteName
let surveyInputDir      = _filestoreRoot @@ "SURVEYS" @@ cleanName
let citInputDir         = _filestoreRoot @@ "CIT" @@ cleanName
let siteWorkInputDir    = _filestoreRoot @@ "SITE_WORK" @@ cleanName
let siteOutputDir       = _outputRoot @@ cleanName



let makeSiteOutputName (fmt:Printf.StringFormat<string->string>) : string = 
    siteOutputDir @@ sprintf fmt cleanName

let makeSiteOutputNamei (fmt:Printf.StringFormat<string->int->string>) (ix:int) : string = 
    siteOutputDir @@ sprintf fmt cleanName ix


Target.Create "Clean" (fun _ -> 
    if Directory.Exists siteOutputDir then 
        Trace.tracefn " --- Clean folder: '%s' ---" siteOutputDir
        Fake.IO.Directory.delete siteOutputDir
    else ()
)

Target.Create "OutputDirectory" (fun _ -> 
    Trace.tracefn " --- Output folder: '%s' ---" siteOutputDir
    maybeCreateDirectory siteOutputDir
)


Target.Create "Cover" (fun _ ->
    let template = _templateRoot @@ "TEMPLATE Samps Cover Sheet.docx"
    let jsonSource = _jsonRoot @@ (sprintf "%s_findreplace.json" cleanName)
    let docname = makeSiteOutputName "%s cover-sheet.docx"
    let pdfname = pathChangeExtension docname "pdf"

    Trace.tracefn " --- Cover sheet for: %s --- " siteName
    
    if File.Exists(jsonSource) then
        let matches = readJsonStringPairs jsonSource    
        DocFindReplace (fun p -> 
            { p with 
                TemplateFile = template
                OutputFile = docname
                Matches = matches
            }) 
    
    
        DocToPdf (fun p -> 
            { p with 
                InputFile = docname
                OutputFile = Some <| pdfname 
            })
    else assertMandatory "CoverSheet failed no json matches"
)


Target.Create "SurveySheet" (fun _ ->
    let infile = Fake.IO.Directory.findFirstMatchingFile "*Sampler survey*.xlsx" surveyInputDir
    let outfile = makeSiteOutputName "%s survey-sheet.pdf" 
    XlsToPdf (fun p -> 
        { p with 
            InputFile = infile
            OutputFile = Some <| outfile
        })
)

let processPhotoFolder (srcPath:string) (destPath:string) : unit = 
    let nameName = System.IO.DirectoryInfo(srcPath).Name
    Trace.tracefn "destPath: %s" destPath
    Trace.tracefn "nameName: %s" nameName
    maybeCreateDirectory destPath 
    multiCopyGlobRename (srcPath, "*.JPG") 
                        (destPath, fun ix -> sprintf "%s %s %03i.jpg" cleanName nameName (ix+1) )
    


Target.Create "SurveyPhotos" (fun _ ->
    let destPath = siteOutputDir @@ "Survey_photos"
    let photoFolders = subdirectoriesWithMatches "*.jpg" surveyInputDir
    List.iter (fun dir -> 
                    Trace.tracefn "Photo dir: %s" dir
                    processPhotoFolder dir destPath) photoFolders
    
    optimizePhotos destPath

    let docName = makeSiteOutputName "%s survey-photos.docx" 
    let pdfName = pathChangeExtension docName "pdf"
    DocPhotos (fun p -> 
        { p with 
            InputPaths = [ destPath ]            
            OutputFile = docName
            ShowFileName = true 
            DocumentTitle = Some <| "Survey Photos"
        })

    DocToPdf (fun p -> 
        { p with 
            InputFile = docName
            OutputFile = Some <| pdfName 
        })
)

// This is optional
Target.Create "SurveyPPT" (fun _ -> 
    match tryFindExactlyOneMatchingFile "*.ppt*" surveyInputDir with
    | Some pptFile -> 
        let outfile = makeSiteOutputName "%s survey-ppt.pdf" 
        Trace.tracefn "Input: %s" pptFile
        PptToPdf (fun p -> 
            { p with 
                InputFile = pptFile
                OutputFile = Some <| outfile
            })
    | None -> assertOptional "No PowerPoint file"       
)

Target.Create "CitCircuitDiagram" (fun _ -> 
    let glob = sprintf "%s*" cleanName
    match tryFindExactlyOneMatchingFile "*Circuit Diagram.pdf" citInputDir with
    | Some srcFile -> 
        let destFile = makeSiteOutputName "%s circuit-diagram.pdf" 
        Fake.IO.Shell.CopyFile destFile srcFile
    | None -> assertOptional "No circuit diagram"
)

// This is optional
Target.Create "CitWorkbook" (fun _ -> 
    match tryFindExactlyOneMatchingFile "*Workbook*.xls*" citInputDir with
    | Some pptFile -> 
        let outfile = makeSiteOutputName "%s cit-workbook.pdf" 
        Trace.tracefn "Input: %s" pptFile
        XlsToPdf (fun p -> 
            { p with 
                InputFile = pptFile
                OutputFile = Some <| outfile
            })
    | None -> assertOptional "No Workbook file"       
)


// Maybe multiple sheets - must find at least 1...
// TODO - potentially "skeletons" to work with multiple files would be nice
Target.Create "InstallSheets" (fun _ ->
    let copySheeti (ix:int) (inputPdf:string) = 
        Trace.tracefn " --- Install sheet %i is: %s --- " (ix+1) inputPdf
        let outfile = makeSiteOutputNamei "%s install-sheet-%03i.pdf" (ix+1)
        if System.IO.File.Exists(inputPdf) then
            Fake.IO.Shell.CopyFile outfile inputPdf 
        else 
            failwithf "InstallSheets - Unbelieveable - glob matches but file does not exist '%s'" inputPdf
    
    match tryFindSomeMatchingFiles "*Replacement Record*.pdf" siteWorkInputDir with
    | Some inputFiles -> List.iteri copySheeti inputFiles 
    | None -> assertMandatory "No Install sheets"
)


Target.Create "InstallPhotos" (fun _ ->

    let photoFolders = subdirectoriesWithMatches "*.jpg" siteWorkInputDir
    if not <| List.isEmpty photoFolders then
        let destPath = siteOutputDir @@ "Install_photos"
        List.iter (fun dir -> 
                        Trace.tracefn "Photo dir: %s" dir
                        processPhotoFolder dir destPath) photoFolders
    
        optimizePhotos destPath

        let docName = makeSiteOutputName "%s install-photos.docx" 
        let pdfName = pathChangeExtension docName "pdf"
        DocPhotos (fun p -> 
            { p with 
                InputPaths = [ destPath ]            
                OutputFile = docName
                ShowFileName = true 
                DocumentTitle = Some <| "Install Photos"
            })

        DocToPdf (fun p -> 
            { p with 
                InputFile = docName
                OutputFile = Some <| pdfName 
            })
    else
        assertOptional "No Install photos"
)

let finalGlobs : string list = 
    [ "*cover-sheet.pdf"
    ; "*survey-sheet.pdf"
    ; "*survey-photos.pdf"
    ; "*cit-workbook.pdf"
    ; "*circuit-diagram.pdf"
    ; "*install-sheet*.pdf" 
    ; "*install-photos.pdf"
    ]


Target.Create "Final" (fun _ ->
    let files:string list= 
        List.collect (fun glob -> findAllMatchingFiles glob siteOutputDir) finalGlobs
    PdfConcat (fun p -> 
        { p with 
            OutputFile = makeSiteOutputName "%s S3820 Sampler Asset Replacement.pdf" })
        files
)



// *** Dependencies ***
"Clean"
    ==> "OutputDirectory"
    ==> "Cover"
    ==> "SurveySheet"
    ==> "SurveyPhotos"
    ==> "SurveyPPT"
    ==> "CitCircuitDiagram"
    ==> "CitWorkbook"
    ==> "InstallSheets"
    ==> "InstallPhotos"
    ==> "Final"

Target.RunOrDefault "Final"

