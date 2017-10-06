

// We have got this path by executing DummyInterop.fsx in F# Interactive
// Clearly this isn't a portable solution, but it allows me to continue development...
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
#load @"DocMake\Tasks\DocFindReplace.fs"
#load @"DocMake\Tasks\DocToPdf.fs"
#load @"DocMake\Tasks\PptToPdf.fs"
#load @"DocMake\Tasks\XlsToPdf.fs"

// Run in PowerShell not fsi:
// PS <path-to-src> ..\packages\FAKE.5.0.0-beta005\tools\FAKE.exe .\BuildScript.fsx Other

// open Microsoft.Office.Interop.Word
open Fake.Core
open DocMake.Tasks.PdfConcat
open DocMake.Tasks.DocFindReplace
open DocMake.Tasks.DocToPdf
open DocMake.Tasks.PptToPdf
open DocMake.Tasks.XlsToPdf

Target.Create "MyBuild" (fun _ ->
    printfn "message from MyBuild target"
)

Target.Create "Other" (fun _ ->
    printfn "Target: Other"
)

Target.Create "Concat" (fun _ -> 
    let (opts:PdfConcatParams->PdfConcatParams) = fun p -> { p with OutputFile = "..\data\output.pdf" }
    let files = [ "..\data\One.pdf"; "..\data\Two.pdf"; "..\data\Three.pdf" ]
    PdfConcat opts files
)

Target.Create "FindReplace" (fun _ -> 
    let opts = fun p -> 
        { p with 
            InputFile = @"D:\coding\fsharp\DocMake\data\findreplace1.docx"
            OutputFile = @"D:\coding\fsharp\DocMake\data\FR-output2.docx" 
            Searches  = [ ("#before", "after") ] }
    DocFindReplace opts
)

Target.Create "DocToPdf" (fun _ -> 
    let (opts:DocToPdfParams->DocToPdfParams) = fun p -> 
        { p with 
            InputFile = @"D:\coding\fsharp\DocMake\data\somedoc.docx" }
    DocToPdf opts
)

Target.Create "XlsToPdf" (fun _ -> 
    let (opts:XlsToPdfParams->XlsToPdfParams) = fun p -> 
        { p with 
            InputFile = @"D:\coding\fsharp\DocMake\data\sheet1.xlsx" }
    XlsToPdf opts
)

Target.Create "PptToPdf" (fun _ -> 
    let (opts:PptToPdfParams->PptToPdfParams) = fun p -> 
        { p with 
            InputFile = @"D:\coding\fsharp\DocMake\data\slides1.pptx" }
    PptToPdf opts
)


Target.RunOrDefault "MyBuild"