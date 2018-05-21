module DocMake.Lib.PptToPdf


// Open at .Interop rather than .PowerPoint then the PowerPoint API has to be qualified
open Microsoft.Office.Interop


open DocMake.Base.Common
open DocMake.Base.OfficeUtils
open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis
open DocMake.Builder.Builders


let private getOutputName (pptDoc:PowerPointDoc) : PowerPointBuild<string> =
    executeIO <| fun () -> 
        System.IO.Path.ChangeExtension(pptDoc.DocumentPath, "pdf")
    


let private process1 (inpath:string) (outpath:string) (quality:DocMakePrintQuality) (app:PowerPoint.Application) : unit = 
    try 
        let prez = app.Presentations.Open(inpath)
        prez.ExportAsFixedFormat (Path = outpath,
                                    FixedFormatType = PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF,
                                    Intent = powerpointPrintQuality quality) 
        prez.Close();
    with
    | ex -> printfn "PptToPdf - Some error occured for %s - '%s'" inpath ex.Message




let pptToPdf (pptDoc:PowerPointDoc) : PowerPointBuild<PdfDoc> =
    buildMonad { 
        let! (app:PowerPoint.Application) = askU ()
        let! outPath = getOutputName pptDoc
        let! quality = asksEnv (fun s -> s.PrintQuality)
        let _ =  process1 pptDoc.DocumentPath outPath quality app
        return (makeDocument outPath)
    }