[<AutoOpen>]
module DocMake.Tasks.DocToPdf

open System.IO
open System.Text.RegularExpressions

// Open at .Interop rather than .Word then the Word API has to be qualified
open Microsoft.Office.Interop

open Fake
open Fake.Core

open DocMake.Base.Common
open DocMake.Base.Office

[<CLIMutable>]
type DocToPdfParams = 
    { 
        InputFile : string
        // If output file is not specified just change extension to .pdf
        OutputFile : string option
        PrintQuality : DocMakePrintQuality
    }

let DocToPdfDefaults = 
    { InputFile = @""
      OutputFile = None
      PrintQuality = PqScreen }


let private getOutputName (opts:DocToPdfParams) : string =
    match opts.OutputFile with
    | None -> System.IO.Path.ChangeExtension(opts.InputFile, "pdf")
    | Some(s) -> s


let private process1 (app:Word.Application) (inpath:string) (outpath:string) (quality:DocMakePrintQuality) : unit = 
    try 
        let doc = app.Documents.Open(FileName = refobj inpath)
        doc.ExportAsFixedFormat (OutputFileName = outpath, 
                                  ExportFormat = Word.WdExportFormat.wdExportFormatPDF,
                                  OptimizeFor = wordPrintQuality quality)
        doc.Close (SaveChanges = refobj false)
    with
    | ex -> printfn "Some error occured - %s - %s" inpath ex.Message




let DocToPdf (setDocToPdfParams: DocToPdfParams -> DocToPdfParams) : unit =
    let options = DocToPdfDefaults |> setDocToPdfParams
    if File.Exists(options.InputFile) 
    then
        let app = new Word.ApplicationClass (Visible = true)
        try 
            process1 app options.InputFile (getOutputName options) options.PrintQuality
        finally 
            app.Quit ()
    else 
        Trace.traceError <| sprintf "DocToPdf --- missing input file"
        failwith "DocToPdf --- missing input file"