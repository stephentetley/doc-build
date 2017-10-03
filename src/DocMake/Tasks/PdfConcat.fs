[<AutoOpen>]
module DocMake.Tasks.PdfConcat

open System.IO
open System.Text.RegularExpressions

// open Fake
open Fake.Core
open Fake.Core.Process
open DocMake.Utils.Common

 

// Concat PDFs
    

[<CLIMutable>]
type PdfConcatParams = 
    { 
        Output : string
        ToolPath : string
        GsOptions : string
    }

// Note input files are supplied as arguments to the top level "Command".
// e.g CscHelper.fs
// Csc : (CscParams -> CscParams) * string list -> unit



// ArchiveHelper.fs includes output file name as a function argument
// DotCover.fs includes output file name in the params
// FscHelper includes output file name in the params

let PdfConcatDefaults = 
    { Output = "concat.pdf"
      ToolPath = "C:\programs\gs\gs9.15\bin\gswin64c.exe"
      GsOptions = @"-dBATCH -dNOPAUSE -q -sDEVICE=pdfwrite -dPDFSETTINGS=/prepress" }


let line1 (opts:PdfConcatParams) : string =
    sprintf "%s %s -sOutputFile=%s" (doubleQuote opts.ToolPath) opts.GsOptions (doubleQuote opts.Output)

let lineK (name:string) : string = sprintf " \"%s\"" name

let unlines (lines: string list) : string = String.concat "\n" lines
let unlinesC (lines: string list) : string = String.concat "^\n" lines

let makeCmd (setPdfConcatParams: PdfConcatParams -> PdfConcatParams) (inputFiles: string list) : string = 
    let parameters = PdfConcatDefaults |> setPdfConcatParams
    let first = line1 parameters
    let rest = List.map lineK inputFiles
    unlinesC <| first :: rest


let private run toolPath command = 
    if 0 <> ExecProcess (fun info -> 
                info.FileName <- toolPath
                info.Arguments <- command) System.TimeSpan.MaxValue
    then failwithf "PdfConcat %s failed." command

let teststring = "TEST_STRING"
    // TODO run as a process...


