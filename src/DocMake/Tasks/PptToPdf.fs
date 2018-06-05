// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

module DocMake.Tasks.PptToPdf

open System.IO
open System.Text.RegularExpressions

open Microsoft.Office
open Microsoft.Office.Interop

open Fake
open Fake.Core

open DocMake.Base.Common
open DocMake.Base.OfficeUtils


[<CLIMutable>]
type PptToPdfParams = 
    { InputFile: string
      // If output file is not specified just change extension to .pdf
      OutputFile: string option
      PrintQuality: DocMakePrintQuality }

let PptToPdfDefaults = 
    { InputFile = @""
      OutputFile = None 
      PrintQuality = PqScreen }


let private getOutputName (opts:PptToPdfParams) : string =
    match opts.OutputFile with
    | None -> System.IO.Path.ChangeExtension(opts.InputFile, "pdf")
    | Some(s) -> s


let private process1 (app:PowerPoint.Application) (inpath:string) (outpath:string) (quality:DocMakePrintQuality) : unit = 
    try 
        let prez = app.Presentations.Open(inpath)
        prez.ExportAsFixedFormat (Path = outpath,
                                    FixedFormatType = PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF,
                                    Intent = powerpointPrintQuality quality) 
        prez.Close();
    with
    | ex -> printfn "PptToPdf - Some error occured for %s - '%s'" inpath ex.Message




let PptToPdf (setPptToPdfParams: PptToPdfParams -> PptToPdfParams) : unit =
    let options = PptToPdfDefaults |> setPptToPdfParams
    if File.Exists(options.InputFile) 
    then
        // This has been leaving a copy of Powerpoint open...
        let app = new PowerPoint.ApplicationClass()
        try 
            app.Visible <- Core.MsoTriState.msoTrue
            process1 app options.InputFile (getOutputName options) options.PrintQuality
            app.Quit ()
        with 
        | ex -> failwithf "PptToPdf - Some error occured for %s - '%s'" options.InputFile ex.Message    
    else 
        Trace.traceError <| sprintf "PptToPdf --- missing input file"
        failwithf "PptToPdf - missing input file '%s'" options.InputFile