[<AutoOpen>]
module DocMake.Tasks.XlsToPdf

open System.IO
open System.Text.RegularExpressions

open Microsoft.Office.Interop

open DocMake.Utils.Office


[<CLIMutable>]
type XlsToPdfParams = 
    { 
        InputFile : string
        // If output file is not specified just change extension to .pdf
        OutputFile : string option
    }

let XlsToPdfDefaults = 
    { InputFile = @""
      OutputFile = None }


let getOutputName (opts:XlsToPdfParams) : string =
    match opts.OutputFile with
    | None -> System.IO.Path.ChangeExtension(opts.InputFile, "pdf")
    | Some(s) -> s


let process1 (app:Excel.Application) (inpath:string) (outpath:string) : unit = 
    try 
        let xls = app.Workbooks.Open(inpath)
        xls.ExportAsFixedFormat (Type=Excel.XlFixedFormatType.xlTypePDF,
                                 Filename=outpath,
                                 IncludeDocProperties=true)
        xls.Close (SaveChanges = false)
    with
    | ex -> printfn "Some error occured - %s - %s" inpath ex.Message




let XlsToPdf (setXlsToPdfParams: XlsToPdfParams -> XlsToPdfParams) : unit =
    let opts = XlsToPdfDefaults |> setXlsToPdfParams
    if File.Exists(opts.InputFile) 
    then
        let app = new Excel.ApplicationClass(Visible = true) 
        try 
            process1 app opts.InputFile (getOutputName opts)
        finally 
            app.Quit ()
    else ()