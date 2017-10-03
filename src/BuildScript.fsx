

// #r "Microsoft.Office.Interop.Word"

#I @"..\packages\FAKE.5.0.0-beta005\tools"
#r @"..\packages\FAKE.5.0.0-beta005\tools\FakeLib.dll"


#load @"DocMake\Utils\Common.fs"
#load @"DocMake\Tasks\PdfConcat.fs"
// #load @"DocMake\Tasks\DocToPdf.fs"

// Run in PowerShell not fsi:
// PS <path-to-src> ..\packages\FAKE.5.0.0-beta005\tools\FAKE.exe .\BuildScript.fsx Other

// open Microsoft.Office.Interop.Word
open Fake.Core
open DocMake.Tasks.PdfConcat

Target.Create "MyBuild" (fun _ ->
    printfn "message from MyBuild target"
)

Target.Create "Other" (fun _ ->
    printfn "Target: Other"
    printfn "%s" teststring
)


Target.RunOrDefault "MyBuild"