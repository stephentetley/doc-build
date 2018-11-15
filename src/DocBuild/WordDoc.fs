// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


module DocBuild.WordDoc


// Open at .Interop rather than .Word then the Word API has to be qualified
open Microsoft.Office.Interop

open DocBuild.PdfDoc

let private rbox (x:'a) : ref<obj> = ref (x :> obj)

let private withWordApp (operation:Word.Application -> 'a) : 'a = 
    let app = new Word.ApplicationClass (Visible = true) :> Word.Application
    let result = operation app
    app.Quit()
    result

type WordExportQuality = 
    | WordForScreen
    | WordForPrint


let private wordPrintQuality (quality:WordExportQuality) : Word.WdExportOptimizeFor = 
    match quality with
    | WordForScreen -> Word.WdExportOptimizeFor.wdExportOptimizeForOnScreen
    | WordForPrint -> Word.WdExportOptimizeFor.wdExportOptimizeForPrint


type WordDoc = 
    val DocPath : string

    new (filePath:string) = 
        { DocPath = filePath }

    member internal v.Body 
        with get() : string = v.DocPath

    member v.ExportAsPdf(quality:WordExportQuality, outFile:string) : PdfDoc = 
        withWordApp <| fun wordApp -> 
            try 
                let doc:(Word.Document) = wordApp.Documents.Open(FileName = rbox v.Body)
                doc.ExportAsFixedFormat (OutputFileName = outFile, 
                                          ExportFormat = Word.WdExportFormat.wdExportFormatPDF,
                                          OptimizeFor = wordPrintQuality quality)
                doc.Close (SaveChanges = rbox false)
                pdfDoc outFile
            with
            | ex -> failwithf "Some error occured - %s - %s" v.Body ex.Message


        member v.ExportAsPdf(quality:WordExportQuality) : PdfDoc =
            let outFile:string = System.IO.Path.ChangeExtension(v.Body, "pdf")
            v.ExportAsPdf(quality= quality, outFile = outFile)

let wordDoc (path:string) : WordDoc = new WordDoc (filePath = path)