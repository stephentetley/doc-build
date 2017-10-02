


#I @"..\packages\FAKE.5.0.0-beta003\tools"
#r @"..\packages\FAKE.5.0.0-beta003\tools\FakeLib.dll"


#load @".\DocMake\Tasks\PdfConcat.fs"

// Run in PowerShell not fsi:
// PS <path-to-src> ..\packages\FAKE.5.0.0-beta003\tools\FAKE.exe .\BuildScript.fsx Other

open Fake.Core
open DocMake.Tasks.PdfConcat

Target.Create "MyBuild" (fun _ ->
    printfn "message from MyBuild target"
)

Target.Create "Other" (fun _ ->
    printfn "Target: Other"
)


Target.RunOrDefault "MyBuild"