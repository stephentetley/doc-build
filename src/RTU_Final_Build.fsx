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
#load @"DocMake\Base\ImageMagick.fs"
#load @"DocMake\Base\Json.fs"
open DocMake.Base.Common

#load @"DocMake\Base\Office.fs"

#load @"DocMake\Tasks\DocFindReplace.fs"
open DocMake.Tasks.DocFindReplace

#load @"DocMake\Tasks\DocPhotos.fs"
open DocMake.Tasks.DocPhotos

#load @"DocMake\Tasks\DocToPdf.fs"
open DocMake.Tasks.DocToPdf

#load @"DocMake\Tasks\PdfConcat.fs"
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
let siteName = environVarOrDefault "sitename" @"CUDWORTH/NO 2 STW"


let cleanName       = safeName siteName
let siteData        = _filestoreRoot @@ cleanName
let siteOutput      = _outputRoot @@ cleanName


let makeSiteOutputName (fmt:Printf.StringFormat<string->string>) : string = 
    siteOutput @@ sprintf fmt cleanName

Target.Create "Clean" (fun _ -> 
    if Directory.Exists(siteOutput) then 
        Trace.tracefn " --- Clean folder: '%s' ---" siteOutput
        Fake.IO.Directory.delete siteOutput
    else ()
)


Target.Create "OutputDirectory" (fun _ -> 
    Trace.tracefn " --- Output folder: '%s' ---" siteOutput
    maybeCreateDirectory(siteOutput)
)

Target.Create "CoverSheet" (fun _ ->
    let template = _templateRoot @@ "MM3x-to-MMIM RTU Cover Sheet.docx"
    let jsonSource = _jsonRoot @@ (sprintf "%s_findreplace.json" cleanName)
    let docname = makeSiteOutputName "%s Cover Sheet.docx"
    Trace.tracefn " --- Cover sheet for: %s --- " siteName
    
    DocFindReplace (fun p -> 
        { p with 
            TemplateFile = template
            OutputFile = docname
            JsonMatchesFile  = jsonSource 
        }) 
    
    let pdfname = makeSiteOutputName "%s Cover Sheet.pdf"
    DocToPdf (fun p -> 
        { p with 
            InputFile = docname
            OutputFile = Some <| pdfname 
        })
)

Target.Create "SurveySheet" (fun _ ->
    let action (infile:string) : unit =
        let outfile = makeSiteOutputName "%s Survey Sheet.pdf" 
        Trace.tracefn " --- Survey sheet is: %s --- " infile
        DocToPdf (fun p -> 
            { p with 
                InputFile = infile
                OutputFile = Some <| outfile
            })

    let inOpt = Fake.IO.Directory.tryFindFirstMatchingFile "* survey.doc" siteData
    match inOpt with
    | Some(infile) -> action infile
    | None -> Trace.tracefn " --- NO SURVEY SHEET --- "
)


Target.Create "SurveyPhotos" (fun _ ->
    let photosPath = siteData @@ "Survey Photos"
    let docname = makeSiteOutputName "%s Survey Photos.docx" 
    let pdfname = makeSiteOutputName "%s Survey Photos.pdf"

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
    else Trace.tracefn " --- NO SURVEY PHOTOS --- "
)

Target.Create "InstallSheet" (fun _ ->
    let infile = Fake.IO.Directory.findFirstMatchingFile "* Site Works*.doc*" siteData
    Trace.tracefn " --- Install sheet is: %s --- " infile
    let outfile = makeSiteOutputName "%s Install Sheet.pdf" 
    if System.IO.File.Exists(infile) then
        DocToPdf (fun p -> 
            { p with 
                InputFile = infile
                OutputFile = Some <| outfile
            })
    else Trace.tracefn " --- NO INSTALL SHEET --- "
)

Target.Create "InstallPhotos" (fun _ ->
    let photosPath = siteData @@ "install photos"
    let docname = makeSiteOutputName "%s Install Photos.docx" 
    let pdfname = makeSiteOutputName "%s Install Photos.pdf"

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
    else Trace.tracefn " --- NO INSTALL PHOTOS --- "
)

let finalGlobs : string list = 
    [ "* Cover Sheet.pdf" ;
      "* Survey Sheet.pdf" ;
      "* Survey Photos.pdf" ;
      "* Install Sheet.pdf" ;
      "* Install Photos.pdf" ]

Target.Create "Final" (fun _ ->
    let get1 (glob:string) : option<string> = 
        Fake.IO.Directory.tryFindFirstMatchingFile glob siteOutput
    let outfile = makeSiteOutputName "%s S3953 RTU Asset Replacement.pdf"
    let files = List.map get1 finalGlobs |> List.choose id
    PdfConcat (fun p ->  { p with OutputFile = outfile })
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
