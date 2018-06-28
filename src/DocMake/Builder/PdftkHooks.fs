// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


module DocMake.Builder.PdftkHooks

open System.IO

open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis


type PdftkHandle = 
    { PdftkExePath: string }


let pdftkBuilderHook (exePath:string) : BuilderHooks<PdftkHandle> = 
    let active = { PdftkExePath = exePath } 
    { InitializeResource = fun _ ->  active
      FinalizeResource = fun _ -> () }



let pdftkRunCommand (getHandle:'res -> PdftkHandle) (command:string) : BuildMonad<'res,unit> = 
    buildMonad { 
        let! toolPath = asksU (getHandle >> fun e -> e.PdftkExePath) 
        do! shellRun toolPath command "Pdftk failed"
    }

