// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


module DocMake.Builder.GhostscriptHooks

open System.IO

open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis



type GsHandle = 
    { GhostscriptExePath: string }


let ghostsciptBuilderHook (exePath:string) : BuilderHooks<GsHandle> = 
    let active = { GhostscriptExePath = exePath } 
    { InitializeResource = fun _ ->  active
      FinalizeResource = fun _ -> () }


let gsRunCommand (getHandle:'res -> GsHandle) (command:string) : BuildMonad<'res,unit> = 
    buildMonad { 
        let! toolPath = asksU (getHandle >> fun e -> e.GhostscriptExePath) 
        do! shellRun toolPath command "GS failed"
    }
