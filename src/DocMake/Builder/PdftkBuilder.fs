// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


module DocMake.Builder.PdftkBuilder

open System.IO
open Microsoft.Office.Interop



open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis


type PdftkEnv = 
    { PdftkExePath: string }

type PdftkBuild<'a> = BuildMonad<PdftkEnv, 'a>


let execPdftkBuild (pathToPdftkExe:string) (ma:PdftkBuild<'a>) : BuildMonad<'res,'a> = 
    let gsEnv = { PdftkExePath = pathToPdftkExe }
    withUserHandle gsEnv (fun _ -> ()) ma

let pdftkRunCommand (command:string) : PdftkBuild<unit> = 
    buildMonad { 
        let! toolPath = asksU (fun (e:PdftkEnv) -> e.PdftkExePath) 
        do! shellRun toolPath command "Pdftk failed"
    }