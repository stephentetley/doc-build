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


#I @"..\packages\Magick.NET-Q8-AnyCPU.7.4.6\lib\net40"
#r @"Magick.NET-Q8-AnyCPU.dll"
open ImageMagick

#I @"..\packages\Newtonsoft.Json.11.0.2\lib\net45"
#r "Newtonsoft.Json"
open Newtonsoft.Json

open System.IO

// FAKE is local to the project file
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
#I @"..\packages\Fake.Core.Target.5.0.0-rc017.237\lib\net46"
#r @"Fake.Core.Target.dll"
#I @"..\packages\Fake.Core.Environment.5.0.0-rc017.237\lib\net46"
#r @"Fake.Core.Environment.dll"
open Fake
open Fake.Core
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
open DocMake.Base.JsonUtils
open DocMake.Base.CopyRename


#load @"DocMake\Tasks\DocFindReplace.fs"
#load @"DocMake\Tasks\DocPhotos.fs"
#load @"DocMake\Tasks\DocToPdf.fs"
#load @"DocMake\Tasks\XlsToPdf.fs"
#load @"DocMake\Tasks\PdfConcat.fs"
open DocMake.Tasks.DocFindReplace
open DocMake.Tasks.DocPhotos
open DocMake.Tasks.DocToPdf
open DocMake.Tasks.XlsToPdf
open DocMake.Tasks.PdfConcat

// NOTE - can generate a batch file to do "many-to-one"
// We usually have many final docs to make (many sites) but the style of 
// Fake is to make one agglomerate out of many parts.
// Generating a batch file that invokes Fake for each site solves this.

let _filestoreRoot  = @"G:\work\Projects\flow2\final-docs\Input\Batch02"
let _outputRoot     = @"G:\work\Projects\flow2\final-docs\output"
let _templateRoot   = @"G:\work\Projects\flow2\final-docs\__Templates"
let _jsonRoot       = @"G:\work\Projects\flow2\final-docs\__Json"

// siteName is an envVar so we can use this build script to build many 
// sites (they all follow the same directory/file structure).
let siteName = Fake.Core.Environment.environVarOrDefault "sitename" @"MISSING"


let cleanName           = safeName siteName
let siteInputDir        = _filestoreRoot @@ cleanName
let siteOutputDir       = _outputRoot @@ cleanName


let makeSiteOutputName (fmt:Printf.StringFormat<string->string>) : string = 
    siteOutputDir @@ sprintf fmt cleanName

Fake.Core.Target.create "Clean" (fun _ -> 
    if Directory.Exists(siteOutputDir) then 
        Trace.tracefn " --- Clean folder: '%s' ---" siteOutputDir
        Fake.IO.Directory.delete siteOutputDir
    else 
        Trace.tracefn " --- Clean --- : folder does not exist '%s' ---" siteOutputDir
)


Fake.Core.Target.create "OutputDirectory" (fun _ -> 
    Trace.tracefn " --- Output folder: '%s' ---" siteOutputDir
    maybeCreateDirectory siteOutputDir 
)



Fake.Core.Target.create "Cover" (fun _ ->
    try
        let template = _templateRoot @@ "FC2 Cover TEMPLATE.docx"
        let jsonSource = _jsonRoot @@ (sprintf "%s_findreplace.json" cleanName)
        let docName = makeSiteOutputName "%s cover-sheet.docx"
        let pdfName = pathChangeExtension docName "pdf"
        Trace.tracefn "Json source: '%s'" jsonSource
        if File.Exists(jsonSource) then
            let matches = readJsonStringPairs jsonSource
            DocFindReplace (fun p -> 
                { p with 
                    TemplateFile = template
                    OutputFile = docName
                    Matches = matches
                }) 
            DocToPdf (fun p -> 
                { p with 
                    InputFile = docName
                    OutputFile = Some <| pdfName 
                })
        else 
            assertMandatory <| sprintf "CoverSheet failed no json matches: %s" jsonSource
    with ex -> assertMandatory <| sprintf "COVER SHEET FAILED\n%s" (ex.ToString())
)


Fake.Core.Target.create "SurveyPhotos" (fun _ ->
    let photosPath = siteInputDir @@ "Survey_Photos"
    let docname = makeSiteOutputName "%s survey-photos.docx" 
    let pdfname = pathChangeExtension docname "pdf"

    if System.IO.Directory.Exists(photosPath) then
        DocPhotos (fun p -> 
            { p with 
                InputPaths = [photosPath]            
                OutputFile = docname
                ShowFileName = true 
            })
        DocToPdf (fun p -> 
            { p with 
                InputFile = docname
                OutputFile = Some <| pdfname 
            })
    else assertOptional "NO INSTALL PHOTOS"
)

Fake.Core.Target.create "ScopeOfWorks" (fun _ ->
    // Note - matching is with globs not regexs. Cannot use [Ss] to match capital or lower s.
    match tryFindExactlyOneMatchingFile "*Scope of Works*.pdf*" siteInputDir with
    | Some source -> 
        let destPath = makeSiteOutputName "%s scope-of-works.pdf"
        optionalCopyFile destPath source
    | None -> assertMandatory "NO SCOPE OF WORKS"
)


// Does not throw error on failure
let copySingletonAction (pattern:string) (srcDir:string) (destPath:string) : unit = 
    match tryFindExactlyOneMatchingFile pattern srcDir with
    | Some source -> 
        optionalCopyFile destPath source
    | None -> assertOptional (sprintf "No match for '%s'" pattern)

Fake.Core.Target.create "Electricals" (fun _ -> 
    // This is optional for all three
    let proc (glob:string, srcDir:string, destPath:string) : unit =
        try 
            copySingletonAction glob srcDir destPath
        with
        | ex -> Trace.tracef "%s" (ex.ToString())

    List.iter proc [ "*Cable Diagram*.pdf",         siteInputDir,   makeSiteOutputName "%s electricals-cable.pdf" 
                   ; "*Termination Drawing*.pdf",   siteInputDir,   makeSiteOutputName "%s electricals-term.pdf" 
                   ; "*Packing list*.pdf",          siteInputDir,   makeSiteOutputName "%s electricals-Zpacking.pdf" 
                   ]
)




Fake.Core.Target.create "InstallSheets" (fun _ ->
    let pdfGen (glob:string) (outputFile:string) (warnMsg:string) : unit = 
        match tryFindExactlyOneMatchingFile glob siteInputDir with
        | Some input -> 
            XlsToPdf (fun p -> 
                { p with 
                    InputFile = input
                    OutputFile = Some <| outputFile
                })
        | None -> assertOptional warnMsg

    let flowOut = makeSiteOutputName "%s install-flow.pdf"
    pdfGen "*Flow*eter*.xls*" flowOut "NO FLOW METER INSTALL SHEET"

    let pressureOut = makeSiteOutputName "%s install-pressure.pdf"
    pdfGen "*Pressure*.xls*" pressureOut "NO PRESSURE SENSOR INSTALL SHEET"
)

Fake.Core.Target.create "WorksPhotos" (fun _ ->
    let photosPath = siteInputDir @@ "Install_Photos"
    let docname = makeSiteOutputName "%s works-photos.docx" 
    let pdfname = pathChangeExtension docname "pdf"

    if System.IO.Directory.Exists(photosPath) then
        DocPhotos (fun p -> 
            { p with 
                InputPaths = [photosPath]            
                OutputFile = docname
                ShowFileName = true 
            })
        DocToPdf (fun p -> 
            { p with 
                InputFile = docname
                OutputFile = Some <| pdfname 
            })
    else assertOptional "NO INSTALL PHOTOS"
)

let finalGlobs : string list = 
    [ "*cover-sheet.pdf" 
    ; "*survey-photos.pdf"
    ; "*scope-of-works.pdf" 
    ; "*electricals-*.pdf" 
    ; "*install-*.pdf"
    ; "*works-photos.pdf" 
    ]

Fake.Core.Target.create "Final" (fun _ ->
    let files:string list= 
        List.collect (fun glob -> findAllMatchingFiles glob siteOutputDir) finalGlobs
    PdfConcat (fun p -> 
        { p with 
            OutputFile = makeSiteOutputName "%s S3402 Flow Confirmation Year 2 Manual.pdf" })
        files
)



// *** Dependencies ***
"Clean"
    ==> "OutputDirectory"

"OutputDirectory"
    ==> "Cover"
    ==> "SurveyPhotos"
    ==> "ScopeOfWorks"
    ==> "Electricals"
    ==> "InstallSheets"
    ==> "WorksPhotos"
    ==> "Final"

// Note seemingly Fake files must end with this...
Fake.Core.Target.runOrDefault "None"
