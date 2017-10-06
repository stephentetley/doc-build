[<AutoOpen>]
module DocMake.Tasks.PptToPdf

open System.IO
open System.Text.RegularExpressions

open Microsoft.Office
open Microsoft.Office.Interop

// open DocMake.Utils.Office


[<CLIMutable>]
type PptToPdfParams = 
    { 
        InputFile : string
        // If output file is not specified just change extension to .pdf
        OutputFile : string option
    }

let PptToPdfDefaults = 
    { InputFile = @""
      OutputFile = None }


let getOutputName (opts:PptToPdfParams) : string =
    match opts.OutputFile with
    | None -> System.IO.Path.ChangeExtension(opts.InputFile, "pdf")
    | Some(s) -> s


let process1 (app:PowerPoint.Application) (inpath:string) (outpath:string) : unit = 
    try 
        let ppt = app.Presentations.Open(inpath)
        ppt.SaveAs(FileName=outpath, 
                   FileFormat=PowerPoint.PpSaveAsFileType.ppSaveAsPDF,
                   EmbedTrueTypeFonts = Core.MsoTriState.msoFalse)
        ppt.Close()
    with
    | ex -> printfn "Some error occured - %s - %s" inpath ex.Message




let PptToPdf (setPptToPdfParams: PptToPdfParams -> PptToPdfParams) : unit =
    let opts = PptToPdfDefaults |> setPptToPdfParams
    if File.Exists(opts.InputFile) 
    then
        let app = new PowerPoint.ApplicationClass()
        try 
            app.Visible <- Core.MsoTriState.msoTrue
            process1 app opts.InputFile (getOutputName opts)
        finally 
            app.Quit ()
    else ()