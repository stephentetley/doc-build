// Run in PowerShell not fsi:
// PS> cd <path-to-src>
// PS> ..\packages\FAKE.5.0.0-beta005\tools\FAKE.exe .\RTU_Final_Build.fsx Dummy

// With params:
// PS> ..\packages\FAKE.5.0.0-beta005\tools\FAKE.exe .\RTU_Final_Build.fsx Final --envar assetname="HELLO/WORLD"

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
#load @"DocMake\Base\JsonUtils.fs"
#load @"DocMake\Base\ImageMagickUtils.fs"
#load @"DocMake\Base\JsonUtils.fs"
#load @"DocMake\Base\OfficeUtils.fs"
#load @"DocMake\Base\CopyRename.fs"
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
// We usually have many final docs to make (many assets) but the style of 
// Fake is to make one agglomerate out of many parts.
// Generating a batch file that invokes Fake for each asset solves this.

let _filestoreRoot  = @"G:\work\Projects\kw_plcar\final_docs\Input-Batch01"
let _outputRoot     = @"G:\work\Projects\kw_plcar\final_docs\output"
let _templateRoot   = @"G:\work\Projects\kw_plcar\final_docs\__Templates"
let _jsonRoot       = @"G:\work\Projects\kw_plcar\final_docs\__Json"

// assetName is an envVar so we can use this build script to build many 
// asset (they all follow the same directory/file structure).
let assetName = environVarOrDefault "assetname" @"MISSING"


let cleanName           = safeName assetName
let assetInputDir       = _filestoreRoot @@ cleanName
let assetOutputDir      = _outputRoot @@ cleanName


let makeAssetOutputName (fmt:Printf.StringFormat<string->string>) : string = 
    assetOutputDir @@ sprintf fmt cleanName

Target.Create "Clean" (fun _ -> 
    if Directory.Exists(assetOutputDir) then 
        Trace.tracefn " --- Clean folder: '%s' ---" assetOutputDir
        Fake.IO.Directory.delete assetOutputDir
    else 
        Trace.tracefn " --- Clean --- : folder does not exist '%s' ---" assetOutputDir
)


Target.Create "OutputDirectory" (fun _ -> 
    Trace.tracefn " --- Output folder: '%s' ---" assetOutputDir
    maybeCreateDirectory assetOutputDir 
)



Target.Create "Cover" (fun _ ->
    try
        let template = _templateRoot @@ "TEMPLATE PLC Cover Sheet.docx"
        let jsonSource = _jsonRoot @@ (sprintf "%s_findreplace.json" cleanName)
        let docName = makeAssetOutputName "%s cover-sheet.docx"
        let pdfName = pathChangeExtension docName "pdf"

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
    with ex -> assertMandatory <| sprintf "COVER SHEET FAILED\n%s" (ex.ToString())
)


// Does not throw error on failure
let copySingletonAction (pattern:string) (srcDir:string) (destPath:string) : unit = 
    match tryFindExactlyOneMatchingFile pattern srcDir with
    | Some source -> 
        optionalCopyFile destPath source
    | None -> assertOptional (sprintf "No match for '%s'" pattern)

Target.Create "CopyPDFs" (fun _ -> 
    // This is optional for all three
    let proc (glob:string, srcDir:string, destPath:string) : unit =
        try 
            copySingletonAction glob srcDir destPath
        with
        | ex -> Trace.tracef "%s" (ex.ToString())

    List.iter proc [ "*FDS V*.pdf",         assetInputDir,   makeAssetOutputName "%s fds.pdf" 
                   ; "*FDS App*.pdf",   assetInputDir,   makeAssetOutputName "%s fds-app.pdf" 
                   ; "*CIRCUIT DIAGRAM*.pdf",          assetInputDir,   makeAssetOutputName "%s circuit-diagram.pdf" 
                   ; "*LOI Screens*.pdf",          assetInputDir,   makeAssetOutputName "%s loi-screens.pdf" 
                   ]
)




let finalGlobs : string list = 
    [ "*cover-sheet.pdf" 
    ; "*fds.pdf" 
    ; "*fds-app.pdf" 
    ; "*circuit-diagram.pdf"
    ; "*loi-screens.pdf" 
    ]

Target.Create "Final" (fun _ ->
    let files:string list= 
        List.collect (fun glob -> findAllMatchingFiles glob assetOutputDir) finalGlobs
    PdfConcat (fun p -> 
        { p with 
            OutputFile = makeAssetOutputName "%s S3821 Asset Replacement.pdf"
            PrintQuality = PdfWhatever 
            })
        files
)



// *** Dependencies ***
"Clean"
    ==> "OutputDirectory"

"OutputDirectory"
    ==> "Cover"
    ==> "CopyPDFs"
    ==> "Final"

// Note seemingly Fake files must end with this...
Target.RunOrDefault "None"
