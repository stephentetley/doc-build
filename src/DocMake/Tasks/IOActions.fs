// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

[<RequireQualifiedAccess>]
module DocMake.Tasks.IOActions

open System.IO

open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis


// These Tasks (Actions) are oblivious to the 'res so we don't need
// to instantiate an API (record-of-functions). 

let clean () : BuildMonad<'res, unit> =
    buildMonad { 
        let! cwd = askWorkingDirectory ()
        if Directory.Exists(cwd) then 
            do printfn " --- Clean folder: '%s' ---" cwd
            do! deleteWorkingDirectory ()
        else 
            do printfn " --- Clean --- : folder does not exist '%s' ---" cwd
    }



let createOutputDirectory () : BuildMonad<'res, unit> =
    buildMonad { 
        let! cwd = asksEnv (fun e -> e.WorkingDirectory)
        do printfn  " --- Output folder: '%s' ---" cwd
        do! createWorkingDirectory ()
    }