[<AutoOpen>]
module DocMake.Tasks.XlsToPdf

open System.IO
open System.Text.RegularExpressions

open Microsoft.Office.Interop

open Fake
open Fake.Core

open DocMake.Base.Common
open DocMake.Base.Office


[<CLIMutable>]
type XlsToPdfParams = 
    { 
        InputFile : string
        // If output file is not specified just change extension to .pdf
        OutputFile : string option
        PrintQuality : DocMakePrintQuality
    }

let XlsToPdfDefaults = 
    { InputFile = @""
      OutputFile = None
      PrintQuality = PqScreen }


let private getOutputName (opts:XlsToPdfParams) : string =
    match opts.OutputFile with
    | None -> System.IO.Path.ChangeExtension(opts.InputFile, "pdf")
    | Some(s) -> s


let private process1 (app:Excel.Application) (inpath:string) (outpath:string) (quality:DocMakePrintQuality) : unit = 
    try 
        let xls = app.Workbooks.Open(inpath)
        xls.ExportAsFixedFormat (Type=Excel.XlFixedFormatType.xlTypePDF,
                                 Filename=outpath,
                                 IncludeDocProperties=true,
                                 Quality = excelPrintQuality quality)
        xls.Close (SaveChanges = false)
    with
    | ex -> printfn "Some error occured - %s - %s" inpath ex.Message




let XlsToPdf (setXlsToPdfParams: XlsToPdfParams -> XlsToPdfParams) : unit =
    let options = XlsToPdfDefaults |> setXlsToPdfParams
    if File.Exists(options.InputFile) 
    then
        let app = new Excel.ApplicationClass(Visible = true) 
        try 
            process1 app options.InputFile (getOutputName options) options.PrintQuality
        finally 
            app.Quit ()
    else 
        Trace.traceError <| sprintf "XlsToPdf --- missing input file"
        failwith "XlsToPdf --- missing input file"