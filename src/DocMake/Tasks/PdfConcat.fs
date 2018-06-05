// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

module DocMake.Tasks.PdfConcat

open System.IO

open Fake.Core

open DocMake.Base.Common

 

// Concat PDFs

// Potentially we could shell out to pdftk, but the space savings don't seem so great:
// > pdftk Input.pdf output Output.pdf compress    

[<CLIMutable>]
type PdfConcatParams = 
    { OutputFile: string
      GhostscriptExePath: string
      PrintQuality: PdfPrintSetting }



// ArchiveHelper.fs includes output file name as a function argument
// DotCover.fs includes output file name in the params
// FscHelper includes output file name in the params



let PdfConcatDefaults = 
    { OutputFile = "concat.pdf"
      GhostscriptExePath = @"C:\programs\gs\gs9.15\bin\gswin64c.exe"
      PrintQuality = PdfScreen }


let private makeGsOptions (quality:PdfPrintSetting) : string =
    match ghostscriptPrintSetting quality with
    | "" -> @"-dBATCH -dNOPAUSE -q -sDEVICE=pdfwrite"
    | ss -> sprintf @"-dBATCH -dNOPAUSE -q -sDEVICE=pdfwrite -dPDFSETTINGS=%s" ss



let private line1 (options:PdfConcatParams) : string =
    sprintf "%s -sOutputFile=%s" 
        (makeGsOptions options.PrintQuality)  (doubleQuote options.OutputFile)

let private lineK (name:string) : string = sprintf " \"%s\"" name

let private unlines (lines: string list) : string = String.concat "\n" lines
let private unlinesC (lines: string list) : string = String.concat "^\n" lines
let private unlinesS (lines: string list) : string = String.concat " " lines

let private makeCmd (parameters: PdfConcatParams) (inputFiles: string list) : string = 
    let first = line1 parameters
    let rest = List.map lineK inputFiles
    unlinesS <| first :: rest

// Run as a process...
let private shellRun toolPath command = 
    let infoF (info:ProcStartInfo) =  
        { info with FileName =  toolPath; Arguments = command }
    if 0 <> Fake.Core.Process.execSimple infoF System.TimeSpan.MaxValue
    then failwithf "PdfConcat %s failed." command



let PdfConcat (setPdfConcatParams: PdfConcatParams -> PdfConcatParams) (inputFiles: string list) : unit =
    let parameters = PdfConcatDefaults |> setPdfConcatParams
    let command = makeCmd parameters inputFiles
    shellRun parameters.GhostscriptExePath command
  


