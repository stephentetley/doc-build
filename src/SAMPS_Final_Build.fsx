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

#I @"..\packages\Newtonsoft.Json.10.0.2\lib\net45"
#r "Newtonsoft.Json"
open Newtonsoft.Json

// FAKE is local to the project file
#I @"..\packages\FAKE.5.0.0-beta005\tools"
#r @"..\packages\FAKE.5.0.0-beta005\tools\FakeLib.dll"

#load @"DocMake\Base\Common.fs"
#load @"DocMake\Base\Json.fs"
#load @"DocMake\Base\Office.fs"
#load @"DocMake\Tasks\PdfConcat.fs"
#load @"DocMake\Tasks\DocPhotos.fs"
#load @"DocMake\Tasks\DocFindReplace.fs"
#load @"DocMake\Tasks\DocToPdf.fs"
#load @"DocMake\Tasks\PptToPdf.fs"
#load @"DocMake\Tasks\UniformRename.fs"
#load @"DocMake\Tasks\XlsToPdf.fs"

open System.IO

open Fake
open Fake.Core
open Fake.Core.Environment
open Fake.Core.Globbing.Operators
open Fake.Core.TargetOperators
// open Fake.Core.Trace
// open Fake opens Fake.EnvironmentHelper     // for (@@) etc.

open DocMake.Base.Common
open DocMake.Tasks.DocFindReplace
open DocMake.Tasks.DocPhotos
open DocMake.Tasks.DocToPdf
open DocMake.Tasks.PdfConcat
open DocMake.Tasks.PptToPdf
open DocMake.Tasks.UniformRename
open DocMake.Tasks.XlsToPdf


let _filestoreRoot  = @"G:\work\Projects\samps\Final_Docs\Jan2018_batch02"
let _outputRoot     = @"G:\work\Projects\samps\Final_Docs\Jan18_OUTPUT"
let _templateRoot   = @"G:\work\Projects\samps\Final_Docs\__Templates"
let _jsonRoot       = @"G:\work\Projects\samps\Final_Docs\__Json"

// siteName is an envVar so we can use this build script to build many 
// sites (they all follow the same directory/file structure).
let siteName = environVarOrDefault "sitename" @"CATTERICK VILLAGE/STW"

let cleanName       = safeName siteName
let siteData        = _filestoreRoot @@ cleanName
let siteOutput      = _outputRoot @@ cleanName



let makeSiteOutputName (fmt:Printf.StringFormat<string->string>) : string = 
    siteOutput @@ sprintf fmt cleanName

let renamePhotos (jpegPath:string) (fmt:Printf.StringFormat<string->int->string>) : unit =
    let mkName = fun i -> sprintf fmt cleanName i
    UniformRename (fun p -> 
        { p with 
            InputFolder = Some <| jpegPath
            MatchPattern = @"\.je?pg$"
            MatchIgnoreCase = true
            MakeName = mkName 
        })

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
    let template = _templateRoot @@ "TEMPLATE Samps Cover Sheet.docx"
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
    let infile = Fake.IO.Directory.findFirstMatchingFile "* Sampler survey.xlsx" (siteData)
    let outfile = makeSiteOutputName "%s Survey Sheet.pdf" 
    XlsToPdf (fun p -> 
        { p with 
            InputFile = infile
            OutputFile = Some <| outfile
        })
)


Target.Create "SurveyPhotos" (fun _ ->
    let inletCopyPath = siteOutput @@ "SurveyPhotos\Inlet"
    maybeCreateDirectory inletCopyPath 
    !! (siteData @@ "Inlet\*.jpg") |> Fake.IO.Shell.Copy inletCopyPath
    renamePhotos inletCopyPath "%s Inlet %03i.jpg"

    let outletCopyPath = siteOutput @@ "SurveyPhotos\Outlet"
    maybeCreateDirectory outletCopyPath
    !! (siteData @@ "Outlet\*.jpg") |> Fake.IO.Shell.Copy outletCopyPath 
    renamePhotos outletCopyPath "%s Outlet %03i.jpg"

    let docname = makeSiteOutputName "%s Survey Photos.docx" 
    DocPhotos (fun p -> 
        { p with 
            InputPaths = [ inletCopyPath; outletCopyPath]            
            OutputFile = docname
            ShowFileName = true 
        })

    let pdfname = makeSiteOutputName "%s Survey Photos.pdf"
    DocToPdf (fun p -> 
        { p with 
            InputFile = docname
            OutputFile = Some <| pdfname 
        })
)

// findFirstMatchingFile is an alternative to unique
Target.Create "SurveyPPT" (fun _ -> 
    let infile = Fake.IO.Directory.findFirstMatchingFile "*.ppt*" (siteData) 
    let outfile = makeSiteOutputName "%s Survey PPT.pdf" 
    Trace.tracef "Input: %s" infile
    PptToPdf (fun p -> 
        { p with 
            InputFile = infile
            OutputFile = Some <| outfile
        })
)

Target.Create "CircuitDiag" (fun _ -> 
    let infile = Fake.IO.Directory.findFirstMatchingFile "* Circuit Diagram.pdf" (siteData) 
    let dest = makeSiteOutputName "%s Circuit Diagram.pdf" 
    Trace.tracef "Input: %s" infile
    Fake.IO.Shell.CopyFile dest infile
)


Target.Create "ElectricalWork" (fun _ ->
    let infile = Fake.IO.Directory.findFirstMatchingFile "* YW Workbook.xls*" (siteData) 
    let outfile = makeSiteOutputName "%s Electrical Worksheet.pdf" 
    XlsToPdf (fun p -> 
        { p with 
            InputFile = infile
            OutputFile = Some <| outfile
        })
)

let finalGlobs : string list = 
    [ "* Cover Sheet.pdf" ;
      "* Survey Sheet.pdf" ;
      "* Survey Photos.pdf" ;
      "* Survey PPT.pdf" ;
      "* Circuit Diagram.pdf" ;
      "* Electrical Worksheet.pdf" ]

//      // For Testing...
//let finalGlobs : string list = 
//    [ "* Install Sheet.pdf" ]

Target.Create "Final" (fun _ ->
    let get1 = fun glob -> Fake.IO.Directory.tryFindFirstMatchingFile glob siteOutput
    let files = List.map get1 finalGlobs |> List.choose id
    PdfConcat (fun p -> 
        { p with 
            OutputFile = makeSiteOutputName "%s S3820 Sampler Asset Replacement.pdf" })
                files
)

Target.Create "Blank" (fun _ ->
    Trace.tracefn "Blank, sitename is %s" siteName
)


// *** Dependencies ***
"Clean"
    ==> "OutputDirectory"
    ==> "CoverSheet"
    ==> "SurveySheet"
    ==> "SurveyPhotos"
    ==> "SurveyPPT"
    ==> "CircuitDiag"
    ==> "ElectricalWork"
    ==> "Final"

Target.RunOrDefault "Blank"

