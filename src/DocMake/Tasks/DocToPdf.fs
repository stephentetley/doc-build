[<AutoOpen>]
module DocMake.Tasks.DocToPdf

open System.IO
open System.Text.RegularExpressions

// Open at .Interop rather than .Word then the Word API has to be qualified
open Microsoft.Office.Interop

open DocMake.Base.Office

[<CLIMutable>]
type DocToPdfParams = 
    { 
        InputFile : string
        // If output file is not specified just change extension to .pdf
        OutputFile : string option
    }

let DocToPdfDefaults = 
    { InputFile = @""
      OutputFile = None }


let private getOutputName (opts:DocToPdfParams) : string =
    match opts.OutputFile with
    | None -> System.IO.Path.ChangeExtension(opts.InputFile, "pdf")
    | Some(s) -> s


let private process1 (app:Word.Application) (inpath:string) (outpath:string) : unit = 
    try 
        let doc = app.Documents.Open(FileName = refobj inpath)
        doc.ExportAsFixedFormat (OutputFileName = outpath, ExportFormat = Word.WdExportFormat.wdExportFormatPDF)
        doc.Close (SaveChanges = refobj false)
    with
    | ex -> printfn "Some error occured - %s - %s" inpath ex.Message




let DocToPdf (setDocToPdfParams: DocToPdfParams -> DocToPdfParams) : unit =
    let opts = DocToPdfDefaults |> setDocToPdfParams
    if File.Exists(opts.InputFile) 
    then
        let app = new Word.ApplicationClass (Visible = true)
        try 
            process1 app opts.InputFile (getOutputName opts)
        finally 
            app.Quit ()
    else ()