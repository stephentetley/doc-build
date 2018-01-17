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
#load @"DocMake\Base\Fake.fs"
#load @"DocMake\Base\ImageMagick.fs"
#load @"DocMake\Base\Json.fs"
#load @"DocMake\Base\Office.fs"
#load @"DocMake\Base\CopyRename.fs"
#load @"DocMake\Tasks\PdfConcat.fs"
#load @"DocMake\Tasks\DocPhotos.fs"
#load @"DocMake\Tasks\DocFindReplace.fs"
#load @"DocMake\Tasks\DocToPdf.fs"
#load @"DocMake\Tasks\PptToPdf.fs"
#load @"DocMake\Tasks\XlsToPdf.fs"


// open Fake.Core.Trace
// open Fake opens Fake.EnvironmentHelper     // for (@@) etc.

open DocMake.Base.Common
open DocMake.Base.Fake
open DocMake.Base.ImageMagick
open DocMake.Base.CopyRename
open DocMake.Tasks.DocFindReplace
open DocMake.Tasks.DocPhotos
open DocMake.Tasks.DocToPdf
open DocMake.Tasks.PdfConcat
open DocMake.Tasks.PptToPdf
open DocMake.Tasks.XlsToPdf


let _filestoreRoot  = @"G:\work\Projects\samps\final-docs\Jan2018_batch02"
let _outputRoot     = @"G:\work\Projects\samps\final-docs\Jan18_OUTPUT"
let _templateRoot   = @"G:\work\Projects\samps\final-docs\__Templates"
let _jsonRoot       = @"G:\work\Projects\samps\final-docs\__Json"

// siteName is an envVar so we can use this build script to build many 
// sites (they all follow the same directory/file structure).
let siteName = environVarOrDefault "sitename" @"CATTERICK VILLAGE/STW"

let cleanName           = safeName siteName
let siteInputDir        = _filestoreRoot @@ cleanName
let siteOutputDir       = _outputRoot @@ cleanName



let makeSiteOutputName (fmt:Printf.StringFormat<string->string>) : string = 
    siteOutputDir @@ sprintf fmt cleanName

let makeSiteOutputNamei (fmt:Printf.StringFormat<string->int->string>) (ix:int) : string = 
    siteOutputDir @@ sprintf fmt cleanName ix


let optimizePhotos (jpegFolderPath:string) : unit =
    let jpegs = !! (jpegFolderPath @@ "*.jpg") |> Seq.toList
    List.iter optimizeForMsWord jpegs
    


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
    
    DocFindReplace (fun p -> 
        { p with 
            TemplateFile = template
            OutputFile = docname
            JsonMatchesFile  = jsonSource
        }) 
    
    
    DocToPdf (fun p -> 
        { p with 
            InputFile = docname
            OutputFile = Some <| pdfname 
        })
)


Target.Create "SurveySheet" (fun _ ->
    let infile = Fake.IO.Directory.findFirstMatchingFile "* Sampler survey.xlsx" (siteInputDir)
    let outfile = makeSiteOutputName "%s survey-sheet.pdf" 
    XlsToPdf (fun p -> 
        { p with 
            InputFile = infile
            OutputFile = Some <| outfile
        })
)


Target.Create "SurveyPhotos" (fun _ ->
    let inletSrcPath    = siteInputDir @@ @"Inlet"
    let inletDestPath   = siteOutputDir @@ @"SurveyPhotos\Inlet"
    maybeCreateDirectory inletDestPath 
    multiCopyGlobRename (inletSrcPath, "*.jpg") 
                        (inletDestPath, fun ix -> sprintf "%s Inlet %03i.jpg" cleanName (ix+1) )
    optimizePhotos inletDestPath

    let outletSrcPath   = siteInputDir @@ "Outlet"
    let outletDestPath  = siteOutputDir @@ @"SurveyPhotos\Outlet"
    maybeCreateDirectory outletDestPath
    multiCopyGlobRename (outletSrcPath, "*.jpg") 
                        (outletDestPath, fun ix -> sprintf "%s Outlet %03i.jpg" cleanName (ix+1) )
    optimizePhotos outletDestPath
    
    let docName = makeSiteOutputName "%s survey-photos.docx" 
    let pdfName = pathChangeExtension docName "pdf"
    DocPhotos (fun p -> 
        { p with 
            InputPaths = [ inletDestPath; outletDestPath ]            
            OutputFile = docName
            ShowFileName = true 
        })

    DocToPdf (fun p -> 
        { p with 
            InputFile = docName
            OutputFile = Some <| pdfName 
        })
)

// This is optional
Target.Create "SurveyPPT" (fun _ -> 
    match tryFindExactlyOneMatchingFile "*.ppt*" (siteInputDir) with
    | Some pptFile -> 
        let outfile = makeSiteOutputName "%s survey-ppt.pdf" 
        Trace.tracef "Input: %s" pptFile
        PptToPdf (fun p -> 
            { p with 
                InputFile = pptFile
                OutputFile = Some <| outfile
            })
    | None -> assertOptional "No PowerPoint file" 
        
)

Target.Create "CircuitDiag" (fun _ -> 
    match tryFindExactlyOneMatchingFile "* Circuit Diagram.pdf" siteInputDir with
    | Some srcFile -> 
        let destFile = makeSiteOutputName "%s circuit-diagram.pdf" 
        Fake.IO.Shell.CopyFile destFile srcFile
    | None -> assertOptional "No circuit diagram"
)

// Not mandatory
Target.Create "ElectricalWork" (fun _ ->
    match tryFindExactlyOneMatchingFile "* YW Workbook.xls*" siteInputDir with
    | Some xlsFile -> 
        let pdfFile = makeSiteOutputName "%s electrical-worksheet.pdf" 
        XlsToPdf (fun p -> 
            { p with 
                InputFile = xlsFile
                OutputFile = Some <| pdfFile
            })
    | None -> assertOptional "No Electrical Work sheets" 
)

Target.Create "BottleMachine" (fun _ -> 
    match tryFindExactlyOneMatchingFile "*ottle*achine*.pdf" siteInputDir with
    | Some srcFile -> 
        let destFile = makeSiteOutputName "%s bottle-machine.pdf" 
        Fake.IO.Shell.CopyFile destFile srcFile
    | None -> assertOptional "No bottle machine certificate"
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
    
    match tryFindSomeMatchingFiles "* Replacement Record.pdf" siteInputDir with
    | Some inputFiles -> List.iteri copySheeti inputFiles 
    | None -> assertMandatory "No Install sheets"
)


let finalGlobs : string list = 
    [ "*cover-sheet.pdf" ;
      "*survey-sheet.pdf" ;
      "*survey-photos.pdf" ;
      "*survey-ppt.pdf" ;
      "*circuit-diagram.pdf" ;
      "*electrical-worksheet.pdf"
      "*bottle-machine.pdf"
      "*install-sheet*.pdf" ]


Target.Create "Final" (fun _ ->
    let files:string list= 
        List.collect (fun glob -> findAllMatchingFiles glob siteOutputDir) finalGlobs
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
    ==> "Cover"
    ==> "SurveySheet"
    ==> "SurveyPhotos"
    ==> "SurveyPPT"
    ==> "CircuitDiag"
    ==> "ElectricalWork"
    ==> "BottleMachine"
    ==> "InstallSheets"
    ==> "Final"

Target.RunOrDefault "Blank"

