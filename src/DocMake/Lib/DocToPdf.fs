module DocMake.Lib.DocToPdf


// Open at .Interop rather than .Word then the Word API has to be qualified
open Microsoft.Office.Interop


open DocMake.Base.Common
open DocMake.Base.OfficeUtils
open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis
open DocMake.Builder.Builders


let private getOutputName (wordDoc:WordDoc) : WordBuild<string> =
    executeIO <| fun () -> 
        System.IO.Path.ChangeExtension(wordDoc.DocumentPath, "pdf")
    

let private process1 (inpath:string) (outpath:string) (quality:DocMakePrintQuality) (app:Word.Application) : unit = 
    try 
        let doc = app.Documents.Open(FileName = refobj inpath)
        doc.ExportAsFixedFormat (OutputFileName = outpath, 
                                  ExportFormat = Word.WdExportFormat.wdExportFormatPDF,
                                  OptimizeFor = wordPrintQuality quality)
        doc.Close (SaveChanges = refobj false)
    with
    | ex -> printfn "Some error occured - %s - %s" inpath ex.Message



let docToPdf (wordDoc:WordDoc) : WordBuild<PdfDoc> =
    buildMonad { 
        let! (app:Word.Application) = askU ()
        let! outPath = getOutputName wordDoc
        let! quality = asksEnv (fun s -> s.PrintQuality)
        let _ =  process1 wordDoc.DocumentPath outPath quality app
        return (makeDocument outPath)
    }