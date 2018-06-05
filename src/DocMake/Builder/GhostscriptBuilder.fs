// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


module DocMake.Builder.GhostscriptBuilder

open System.IO
open Microsoft.Office.Interop



open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis


type GsEnv = 
    { GhostscriptExePath: string }

type GsBuild<'a> = BuildMonad<GsEnv, 'a>


let execGsBuild (pathToGsExe:string) (ma:GsBuild<'a>) : BuildMonad<'res,'a> = 
    let gsEnv = { GhostscriptExePath = pathToGsExe }
    withUserHandle gsEnv (fun _ -> ()) ma


let gsRunCommand (command:string) : GsBuild<unit> = 
    buildMonad { 
        let! toolPath = asksU (fun (e:GsEnv) -> e.GhostscriptExePath) 
        do! shellRun toolPath command "GS failed"
    }
