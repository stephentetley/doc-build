[<AutoOpen>]
module DocMake.Tasks.PdfConcat

open System.IO
open System.Text.RegularExpressions

// open Fake - intellisense not working properly, but we can run this 
// with Fake.exe in PowerShell
open Fake.Core
open Fake.Core.Process

open DocMake.Base.Common

 

// Concat PDFs

// Potentially we could shell out to pdftk, but the space savings don't seem so great:
// > pdftk Input.pdf output Output.pdf compress    

[<CLIMutable>]
type PdfConcatParams = 
    { 
        OutputFile : string
        AppPath : string
        GsOptions : string
    }

// Note input files are supplied as arguments to the top level "Command".
// e.g CscHelper.fs
// Csc : (CscParams -> CscParams) * string list -> unit



// ArchiveHelper.fs includes output file name as a function argument
// DotCover.fs includes output file name in the params
// FscHelper includes output file name in the params


// TODO - GsOptions should be more user friendly
let PdfConcatDefaults = 
    { OutputFile = "concat.pdf"
      AppPath = @"C:\programs\gs\gs9.15\bin\gswin64c.exe"
      GsOptions = @"-dBATCH -dNOPAUSE -q -sDEVICE=pdfwrite -dPDFSETTINGS=/screen" }


let private line1 (opts:PdfConcatParams) : string =
    sprintf "%s -sOutputFile=%s" opts.GsOptions (doubleQuote opts.OutputFile)

let private lineK (name:string) : string = sprintf " \"%s\"" name

let private unlines (lines: string list) : string = String.concat "\n" lines
let private unlinesC (lines: string list) : string = String.concat "^\n" lines
let private unlinesS (lines: string list) : string = String.concat " " lines

let private makeCmd (parameters: PdfConcatParams) (inputFiles: string list) : string = 
    let first = line1 parameters
    let rest = List.map lineK inputFiles
    unlinesS <| first :: rest

// Run as a process...
let private run toolPath command = 
    if 0 <> ExecProcess (fun info -> 
                info.FileName <- toolPath
                info.Arguments <- command) System.TimeSpan.MaxValue
    then failwithf "PdfConcat %s failed." command



let PdfConcat (setPdfConcatParams: PdfConcatParams -> PdfConcatParams) (inputFiles: string list) : unit =
    let parameters = PdfConcatDefaults |> setPdfConcatParams
    let command = makeCmd parameters inputFiles
    run parameters.AppPath command
  


