// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


namespace DocBuild


[<AutoOpen>]
module WordDoc = 


    // Open at .Interop rather than .Word then the Word API has to be qualified
    open Microsoft.Office.Interop

    open DocBuild.Internal.CommonUtils
    open DocBuild.Internal.WordUtils
    open DocBuild



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
        val private SourcePath : string
        val private TempPath : string

        new (filePath:string) = 
            { SourcePath = filePath
            ; TempPath = getTempFileName filePath }

        member internal v.TempFile
            with get() : string = 
                if System.IO.File.Exists(v.TempPath) then
                    v.TempPath
                else
                    System.IO.File.Copy(v.SourcePath, v.TempPath)
                    v.TempPath
    
        member internal v.Updated 
            with get() : bool = System.IO.File.Exists(v.TempPath)
            

        member v.ExportAsPdf( quality:WordExportQuality
                            , outFile:string ) : Document = 
            // Don't make a temp file if we don't have to
            let srcFile = if v.Updated then v.TempPath else v.SourcePath
            withWordApp <| fun app -> 
                try 
                    let doc:(Word.Document) = app.Documents.Open(FileName = rbox srcFile)
                    doc.ExportAsFixedFormat (OutputFileName = outFile, 
                                              ExportFormat = Word.WdExportFormat.wdExportFormatPDF,
                                              OptimizeFor = wordExportQuality quality)
                    doc.Close (SaveChanges = rbox false)
                    pdfDoc outFile
                with
                | ex -> failwithf "Some error occured - %s - %s" srcFile ex.Message

        member v.ExportAsPdf(quality:WordExportQuality) : Document =
            // Don't make a temp file if we don't have to
            let srcFile = if v.Updated then v.TempPath else v.SourcePath
            let outFile:string = System.IO.Path.ChangeExtension(srcFile, "pdf")
            v.ExportAsPdf(quality= quality, outFile = outFile)


        member v.SaveAs(outputPath: string) : unit = 
            if v.Updated then 
                System.IO.File.Move(v.TempPath, outputPath)
            else
                System.IO.File.Copy(v.SourcePath, outputPath)


        member v.FindReplace(searches:SearchList) : WordDoc = 
            withWordApp <| fun app -> 
                let tempFile = v.TempFile
                wordFindReplace app tempFile None searches
            v


    let wordDoc (path:string) : WordDoc = new WordDoc (filePath = path)

