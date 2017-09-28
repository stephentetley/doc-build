
#I @"..\packages\FAKE.5.0.0-beta003\tools"
#r "FakeLib.dll"

// Run in PowerShell not fsi:
// PS <path-to-src> ..\packages\FAKE.5.0.0-beta003\tools\FAKE.exe .\BuildScript.fsx Other

open Fake.Core

Target.Create "MyBuild" (fun _ ->
    printfn "message from MyBuild target"
)

Target.Create "Other" (fun _ ->
    printfn "Target: Other"
)


Target.RunOrDefault "MyBuild"