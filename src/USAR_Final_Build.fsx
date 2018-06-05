// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

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

let _filestoreRoot  = @"G:\work\Projects\usar\Final_Docs\Jan2018_INPUT_batch01"
let _outputRoot     = @"G:\work\Projects\usar\Final_Docs\output"
let _templateRoot   = @"G:\work\Projects\usar\Final_Docs\__Templates"
let _jsonRoot       = @"G:\work\Projects\usar\Final_Docs\__Json"

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

Fake.Core.Target.create "CoverSheet" (fun _ ->
    let template = _templateRoot @@ "USAR Cover Sheet.docx"
    let jsonSource = _jsonRoot @@ (sprintf "%s_findreplace.json" cleanName)
    let docname = makeSiteOutputName "%s Cover Sheet.docx"
    Trace.tracefn " --- Cover sheet for: %s --- " siteName
    
    if File.Exists(jsonSource) then
        let matches = readJsonStringPairs jsonSource
        DocFindReplace (fun p -> 
            { p with 
                TemplateFile = template
                OutputFile = docname
                Matches = matches 
            }) 
    
        let pdfname = makeSiteOutputName "%s Cover Sheet.pdf"
        DocToPdf (fun p -> 
            { p with 
                InputFile = docname
                OutputFile = Some <| pdfname 
            })
    else assertMandatory "CoverSheet failed no json matches"
)

// All file are created in the siteOutputDir...
let docToPdfAction (message:string) (infile:string) : unit =
    let outfile = pathChangeExtension (pathChangeDirectory infile siteOutputDir) "pdf"
    Trace.trace message
    DocToPdf (fun p -> 
        { p with 
            InputFile = infile
            OutputFile = Some <| outfile
        })


// Multiple survey sheets
Fake.Core.Target.create "SurveySheets" (fun _ ->
    // Note - matching is with globs not regexs. Cannot use [Ss] to match capital or lower s.
    match tryFindSomeMatchingFiles "*urvey.doc*" siteInputDir with
    | Some inputs -> 
        List.iter (fun file -> docToPdfAction (sprintf "Survey: %s" file) file) inputs
    | None -> 
        failwith " --- NO SURVEY SHEETS --- "
)



Fake.Core.Target.create "InstallSheets" (fun _ ->
    match tryFindSomeMatchingFiles "*nstall.doc*" siteInputDir with
    | Some inputs -> 
        List.iter (fun file -> docToPdfAction (sprintf "Install: %s" file) file) inputs
    | None -> 
        failwith " --- NO INSTALL SHEETS --- "
)



let finalGlobs : string list = 
    [ "* Cover Sheet.pdf" ;
      "*Survey*.pdf" ;
      "*Install*.pdf" ]

Fake.Core.Target.create "Final" (fun _ ->
    let files:string list= 
        List.collect (fun glob -> findAllMatchingFiles glob siteOutputDir) finalGlobs
    PdfConcat (fun p -> 
        { p with 
            OutputFile = makeSiteOutputName "%s S3820 Ultrasonic Asset Replacement.pdf" })
        files
)

// *** Dummy cases

Fake.Core.Target.create "None" (fun _ ->
    printfn "None"
)

// *** Dependencies ***
"Clean"
    ==> "OutputDirectory"
    ==> "CoverSheet"
    ==> "SurveySheets"
    ==> "InstallSheets"
    ==> "Final"

// Note seemingly Fake files must end with this...
Fake.Core.Target.runOrDefault "None"
