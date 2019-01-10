// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Office


[<RequireQualifiedAccess>]
module WordFile = 

    open System.IO

    // Open at .Interop rather than .Word then the Word API has to be qualified
    open Microsoft.Office.Interop

    open DocBuild.Base.Common
    open DocBuild.Base.Document
    open DocBuild.Office.Internal
    open DocBuild.Office
    open DocBuild.Office.OfficeMonad


    let private wordExportQuality (quality:PrintQuality) : Word.WdExportOptimizeFor = 
        match quality with
        | PqScreen -> Word.WdExportOptimizeFor.wdExportOptimizeForOnScreen
        | PqPrint -> Word.WdExportOptimizeFor.wdExportOptimizeForPrint


    let exportPdfAs (src:WordFile) 
                    (quality:PrintQuality) 
                    (outputFile:string) : OfficeMonad<PdfFile> = 
        officeMonad { 
            let! _ = execWord <| fun app -> 
                        wordExportAsPdf app src.Path (wordExportQuality quality) outputFile
            let! pdf = liftDocMonad (pdfFile outputFile)
            return pdf
        }

    let exportPdf (src:WordFile) (quality:PrintQuality) : OfficeMonad<PdfFile> = 
        let outputFile = Path.ChangeExtension(src.Path, "docx")
        exportPdfAs src quality outputFile




    let findReplaceAs (src:WordFile) (searches:SearchList) (outputFile:string) : OfficeMonad<WordFile> = 
        officeMonad { 
            let! _ = execWord <| fun app -> 
                        wordFindReplace app src.Path outputFile searches
                        Ok ()
            let! docx = liftDocMonad (wordFile outputFile)
            return docx
        }



    let findReplace (src:WordFile) (searches:SearchList)  : OfficeMonad<WordFile> = 
        findReplaceAs src searches src.NextTempName 
