// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Office.WordDoc


[<AutoOpen>]
module WordDoc = 


    // Open at .Interop rather than .Word then the Word API has to be qualified
    open Microsoft.Office.Interop
    open DocBuild.Base
    open DocBuild.Base.Document
    open DocBuild.Office.MsoWord




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


    type WordFile = 
        val private WordDoc : Document

        new (filePath:string) = 
            { WordDoc = new Document(filePath = filePath) }



        member x.ExportAsPdf( quality:WordExportQuality
                            , outFile:string ) : unit = 
            // Don't make a temp file if we don't have to
            let srcFile = x.WordDoc.TempFile
            withWordApp <| fun app -> 
                try 
                    let doc:(Word.Document) = app.Documents.Open(FileName = rbox srcFile)
                    doc.ExportAsFixedFormat (OutputFileName = outFile, 
                                              ExportFormat = Word.WdExportFormat.wdExportFormatPDF,
                                              OptimizeFor = wordExportQuality quality)
                    doc.Close (SaveChanges = rbox false)
                with
                | ex -> failwithf "Some error occured - %s - %s" srcFile ex.Message

        member x.ExportAsPdf(quality:WordExportQuality) : unit =
            // Don't make a temp file if we don't have to
            let srcFile = x.WordDoc.TempFile
            let outFile:string = System.IO.Path.ChangeExtension(srcFile, "pdf")
            x.ExportAsPdf(quality= quality, outFile = outFile)


        member x.SaveAs(outputPath: string) : unit = 
            x.SaveAs outputPath


        member x.FindReplace(searches:SearchList) : unit = 
            withWordApp <| fun app -> 
                let tempFile = x.WordDoc.TempFile
                wordFindReplace app tempFile None searches



    let wordFile (path:string) : WordFile = new WordFile (filePath = path)

