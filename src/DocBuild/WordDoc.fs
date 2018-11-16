// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


module DocBuild.WordDoc


// Open at .Interop rather than .Word then the Word API has to be qualified
open Microsoft.Office.Interop

open DocBuild.Internal.CommonUtils
open DocBuild.PdfDoc


let private withWordApp (operation:Word.Application -> 'a) : 'a = 
    let app = new Word.ApplicationClass (Visible = true) :> Word.Application
    let result = operation app
    app.Quit()
    result

type WordExportQuality = 
    | WordForScreen
    | WordForPrint


let private wordExportQuality (quality:WordExportQuality) : Word.WdExportOptimizeFor = 
    match quality with
    | WordForScreen -> Word.WdExportOptimizeFor.wdExportOptimizeForOnScreen
    | WordForPrint -> Word.WdExportOptimizeFor.wdExportOptimizeForPrint

/// TODO WordDoc should be updateable (e.g. by Find/Replace)
/// But it would be nice not to generate a temp file if all
/// we are doing is exporting to PDF (with no modications).


type WordDoc = 
    val private DocPath : string

    new (filePath:string) = 
        { DocPath = filePath }

    member internal v.Body 
        with get() : string = v.DocPath

    member v.ExportAsPdf(quality:WordExportQuality, outFile:string) : PdfDoc = 
        withWordApp <| fun app -> 
            try 
                let doc:(Word.Document) = app.Documents.Open(FileName = rbox v.Body)
                doc.ExportAsFixedFormat (OutputFileName = outFile, 
                                          ExportFormat = Word.WdExportFormat.wdExportFormatPDF,
                                          OptimizeFor = wordExportQuality quality)
                doc.Close (SaveChanges = rbox false)
                pdfDoc outFile
            with
            | ex -> failwithf "Some error occured - %s - %s" v.Body ex.Message

    member v.ExportAsPdf(quality:WordExportQuality) : PdfDoc =
        let outFile:string = System.IO.Path.ChangeExtension(v.Body, "pdf")
        v.ExportAsPdf(quality= quality, outFile = outFile)

let wordDoc (path:string) : WordDoc = new WordDoc (filePath = path)

