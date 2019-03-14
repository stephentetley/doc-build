// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Office


module MarkdownWordPdf = 

    open System.IO

    // Open at .Interop rather than .Word then the Word API has to be qualified
    open Microsoft.Office.Interop


    open DocBuild.Base
    open DocBuild.Base.DocMonad
    open DocBuild.Document.Markdown
    open DocBuild.Office

    // ************************************************************************
    // Export to Pdf with Pandoc (and TeX)


    let markdownToWordToPdfAs (customStyles:WordDoc option) 
                              (quality:PrintQuality)
                              (outputAbsPath:string) 
                              (src:MarkdownDoc) : DocMonad<#WordDocument.HasWordHandle,PdfDoc> =
        docMonad { 
            let docAbsPath = Path.ChangeExtension(outputAbsPath, "docx")
            let! doc = markdownToWordAs customStyles docAbsPath src
            return! WordDocument.exportPdfAs quality outputAbsPath doc
         }


    let markdownToWordToPdf (customStyles:WordDoc option) 
                            (quality:PrintQuality)                  
                            (src:MarkdownDoc)  : DocMonad<#WordDocument.HasWordHandle,PdfDoc> =
        let outputFile = Path.ChangeExtension(src.LocalPath, "pdf")
        markdownToWordToPdfAs customStyles quality outputFile src


