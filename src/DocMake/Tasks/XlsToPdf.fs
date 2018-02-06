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
    { InputFile: string
      // If output file is not specified just change extension to .pdf
      OutputFile: string option
      PrintQuality: DocMakePrintQuality
      FitWidth: bool }

let XlsToPdfDefaults = 
    { InputFile = @""
      OutputFile = None
      PrintQuality = PqScreen
      FitWidth = true }


let private getOutputName (opts:XlsToPdfParams) : string =
    match opts.OutputFile with
    | None -> System.IO.Path.ChangeExtension(opts.InputFile, "pdf")
    | Some(s) -> s


let private process1 (app:Excel.Application) (inpath:string) (outpath:string) (quality:DocMakePrintQuality) (fitWidth:bool) : unit = 
    try 
        let workbook : Excel.Workbook = app.Workbooks.Open(inpath)
        if fitWidth then 
            workbook.Sheets 
                |> Seq.cast<Excel.Worksheet>
                |> Seq.iter (fun (sheet:Excel.Worksheet) -> 
                    sheet.PageSetup.Zoom <- false
                    sheet.PageSetup.FitToPagesWide <- 1)
        else ()

        workbook.ExportAsFixedFormat (Type=Excel.XlFixedFormatType.xlTypePDF,
                                         Filename=outpath,
                                         IncludeDocProperties=true,
                                         Quality = excelPrintQuality quality
                                         )
        workbook.Close (SaveChanges = false)
    with
    | ex -> Trace.traceError (sprintf "Some error occured - %s - %s" inpath ex.Message)




let XlsToPdf (setXlsToPdfParams: XlsToPdfParams -> XlsToPdfParams) : unit =
    let options = XlsToPdfDefaults |> setXlsToPdfParams
    if File.Exists(options.InputFile) 
    then
        let app = new Excel.ApplicationClass(Visible = true) 
        try 
            process1 app options.InputFile (getOutputName options) options.PrintQuality options.FitWidth
        finally 
            app.Quit ()
    else 
        Trace.traceError <| sprintf "XlsToPdf --- missing input file"
        failwith "XlsToPdf --- missing input file"