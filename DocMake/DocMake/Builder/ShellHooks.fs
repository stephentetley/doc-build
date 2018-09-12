// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


module DocMake.Builder.ShellHooks

open System.IO

open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis



// *************************************
// Ghostscript

type GsHandle = 
    { GhostscriptExePath: string }



let gsRunCommand (getHandle:'res -> GsHandle) (command:string) : BuildMonad<'res,unit> = 
    buildMonad { 
        let! toolPath = asksU (getHandle >> fun e -> e.GhostscriptExePath) 
        do! shellRun toolPath command "GS failed"
    }



// *************************************
// Pdftk

type PdftkHandle = 
    { PdftkExePath: string }



let pdftkRunCommand (getHandle:'res -> PdftkHandle) (command:string) : BuildMonad<'res,unit> = 
    buildMonad { 
        let! toolPath = asksU (getHandle >> fun e -> e.PdftkExePath) 
        do! shellRun toolPath command "Pdftk failed"
    }


// *************************************
// Pandoc


type PandocHandle = 
    { PandocExePath: string }


let pandocRunCommand (getHandle:'res -> PandocHandle) (command:string) : BuildMonad<'res,unit> = 
    buildMonad { 
        let! toolPath = asksU (getHandle >> fun e -> e.PandocExePath) 
        do! shellRun toolPath command "Pandoc failed"
    }

