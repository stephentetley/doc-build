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

// FAKE is local to the project file
#I @"..\packages\FAKE.5.0.0-beta005\tools"
#r @"..\packages\FAKE.5.0.0-beta005\tools\FakeLib.dll"

#load @"DocMake\Utils\Common.fs"
#load @"DocMake\Utils\Office.fs"
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
// open Fake.Core.Trace
// open Fake opens Fake.EnvironmentHelper     // for (@@) etc.

open DocMake.Utils.Common
open DocMake.Tasks.DocFindReplace
open DocMake.Tasks.DocPhotos
open DocMake.Tasks.DocToPdf
open DocMake.Tasks.PdfConcat
open DocMake.Tasks.PptToPdf
open DocMake.Tasks.UniformRename
open DocMake.Tasks.XlsToPdf

let filestoreRoot = @"G:\work\DocMake_TEST"
let templateRoot = @"G:\work\DocMake_TEST\__Templates"
let siteName = environVarOrDefault "sitename" @"BEDALE/STW"
let saiNumber = @"SAI00001647"

let cleanName = safeName siteName
let sitePath = System.IO.Path.Combine(filestoreRoot,cleanName)
let outputRoot = sitePath @@ "__DM_OUTPUT"

let relativeToTemplateRoot (suffix:string) : string = 
    System.IO.Path.Combine(templateRoot, suffix)

let relativeToSite (suffix:string) : string = 
    System.IO.Path.Combine(sitePath, suffix)

let makeSiteOutputName (fmt:Printf.StringFormat<string->string>) : string = 
    outputRoot @@ sprintf fmt cleanName

let renamePhotos (jpegPath:string) (fmt:Printf.StringFormat<string->int->string>) : unit =
    let mkName = fun i -> sprintf fmt cleanName i
    UniformRename (fun p -> 
        { p with 
            InputFolder = jpegPath
            MatchPattern = @"\.je?pg$"
            MatchIgnoreCase = true
            MakeName = mkName 
        })


Target.Create "OutputDirectory" (fun _ -> 
    Trace.tracefn " --- Output folder: '%s' ---" outputRoot
    maybeCreateDirectory(outputRoot)
)

Target.Create "CoverSheet" (fun _ ->
    let template = relativeToTemplateRoot "TEMPLATE Samps Cover Sheet.docx"
    let docname = makeSiteOutputName "%s Cover Sheet.docx"
    Trace.tracefn " --- Cover sheet for: %s --- " siteName
    
    DocFindReplace (fun p -> 
        { p with 
            InputFile = template
            OutputFile = docname
            Searches  = [ ("#SITENAME", siteName);
                          ("#SAINUM", saiNumber) ] 
        }) 
    
    let pdfname = makeSiteOutputName "%s Cover Sheet.pdf"
    DocToPdf (fun p -> 
        { p with 
            InputFile = docname
            OutputFile = Some <| pdfname 
        })
)

Target.Create "SurveyPPT" (fun _ -> 
    let infile = !! (relativeToSite @"1_Survey\*.pptx") |> unique
    let outfile = makeSiteOutputName "%s Survey PPT.pdf" 
    PptToPdf (fun p -> 
        { p with 
            InputFile = infile
            OutputFile = Some <| outfile
        })
)

Target.Create "SurveySheet" (fun _ ->
    let infile = !! (relativeToSite @"1_Survey\*Sampler survey.xlsx") |> unique
    let outfile = makeSiteOutputName "%s Survey Sheet.pdf" 
    XlsToPdf (fun p -> 
        { p with 
            InputFile = infile
            OutputFile = Some <| outfile
        })
)

Target.Create "InstallSheet" (fun _ ->
    let infile = !! (relativeToSite @"2_Site_works\* Wookbook.xls*") |> unique
    let outfile = makeSiteOutputName "%s Install Sheet.pdf" 
    XlsToPdf (fun p -> 
        { p with 
            InputFile = infile
            OutputFile = Some <| outfile
        })
)

Target.Create "SurveyPhotos" (fun _ ->
    let inletpath = (outputRoot @@ "SurveyPhotos\Inlet")
    maybeCreateDirectory inletpath 
    !! (relativeToSite "1_Survey\Inlet\*.jpg") |> FileHelper.Copy inletpath
    renamePhotos inletpath "%s Inlet %03i.jpg"

    let outletpath = (outputRoot @@ "SurveyPhotos\Outlet")
    maybeCreateDirectory outletpath
    !! (relativeToSite "1_Survey\Outlet\*.jpg") |> FileHelper.Copy outletpath 
    renamePhotos outletpath "%s Outlet %03i.jpg"

    let docname = makeSiteOutputName "%s Survey Photos.docx" 
    DocPhotos (fun p -> 
        { p with 
            InputPaths = [ inletpath; outletpath]            
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


// ----------------------------------------------------------------------
// ----------------------------------------------------------------------
// ----------------------------------------------------------------------
// ----------------------------------------------------------------------
// ----- OLD -----
// The functions below confused targets with tasks...

// TODO - Should be a parametric target, because there are two photos folders
Target.Create "RenamePhotos" (fun _ -> 
    let shouldBeParam = @"1_Survey\Inlet"
    let jpegPath = System.IO.Path.Combine(sitePath, shouldBeParam)
    let mkName = fun i -> sprintf "%s Inlet %03i.jpg" cleanName i
    let (opts: UniformRenameParams -> UniformRenameParams) = fun p -> 
        { p with 
            InputFolder = jpegPath
            MatchPattern = @"\.jpg$"
            MatchIgnoreCase = true
            MakeName = mkName }
    UniformRename opts
)


Target.Create "PhotoDoc" (fun _ -> 
    let (opts: DocPhotosParams -> DocPhotosParams) = fun p -> 
        { p with 
            InputPaths = [ relativeToSite @"1_Survey\Inlet";
                            relativeToSite @"1_Survey\Outlet" ]            
            OutputFile = outputRoot @@ (sprintf "%s Survey Photos.docx" cleanName)
            ShowFileName = true }
    DocPhotos opts
)


// Is this really a task or should it be a function?
// Answer - it is a task not a target...
Target.Create "FileCopy" (fun _ -> 
    let srcfile = sitePath @@ "2_Site_works" @@ "Bedale STW Circuit Diagram.pdf"
    let destfile = makeSiteOutputName "%s Circuit Diagram.pdf"
    File.Copy(srcfile,destfile)
)

Target.Create "ConcatFinal" (fun _ ->
    let (opts:PdfConcatParams->PdfConcatParams) = fun p -> 
        { p with 
            OutputFile = makeSiteOutputName "%s UWW Samplers OM Manual.pdf" }
    let files = [ "..\data\One.pdf"; "..\data\Two.pdf"; "..\data\Three.pdf" ]
    PdfConcat opts files
)

Target.Create "Dummy" (fun _ ->
    Trace.tracefn "Dummy, sitename is %s" siteName
)

Target.RunOrDefault "Dummy"