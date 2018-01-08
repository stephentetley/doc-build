[<AutoOpen>]
module DocMake.Tasks.PptToPdf

open System.IO
open System.Text.RegularExpressions

open Microsoft.Office
open Microsoft.Office.Interop

// open DocMake.Base.Office


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


let private getOutputName (opts:PptToPdfParams) : string =
    match opts.OutputFile with
    | None -> System.IO.Path.ChangeExtension(opts.InputFile, "pdf")
    | Some(s) -> s


let private process1 (app:PowerPoint.Application) (inpath:string) (outpath:string) : unit = 
    try 
        // File already exists
        let prez = app.Presentations.Open(inpath)
        prez.SaveAs(FileName=outpath, 
                    FileFormat=PowerPoint.PpSaveAsFileType.ppSaveAsPDF,
                    EmbedTrueTypeFonts = Core.MsoTriState.msoFalse)
        prez.Close();
    with
    | ex -> printfn "PptToPdf - Some error occured for %s - '%s'" inpath ex.Message




let PptToPdf (setPptToPdfParams: PptToPdfParams -> PptToPdfParams) : unit =
    let opts = PptToPdfDefaults |> setPptToPdfParams
    if File.Exists(opts.InputFile) 
    then
        // This has been leaving a copy of Powerpoint open...
        let app = new PowerPoint.ApplicationClass()
        try 
            app.Visible <- Core.MsoTriState.msoTrue
            process1 app opts.InputFile (getOutputName opts)
            app.Quit ()
        with 
        | ex -> printfn "PptToPdf - Some error occured for %s - '%s'" opts.InputFile ex.Message    
    else ()