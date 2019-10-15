// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Office


module PandocWordShim = 

    open System.IO

    // Open Office at .Interop rather than .Word then the Word API 
    // has to be qualified.
    open Microsoft.Office.Interop

    open MarkdownDoc.Markdown

    open DocBuild.Base
    open DocBuild.Base.DocMonad
    open DocBuild.Document.Markdown
    open DocBuild.Extra
    
    open DocBuild.Office

    let asksCustomStyles () : DocMonad< WordDoc option, 'res> = 
        docMonad {
            match! asks (fun env -> env.PandocOpts.CustomStylesDocx) with 
            | None -> return None 
            | Some fileName -> return! getIncludeWordDoc fileName |>> Some
        }


    // ************************************************************************
    // Export to Pdf with Pandoc and MS Word

    /// This uses MS Word as to render in intermediate docx file to Pdf.
    let markdownToPdfAs (outputPdfName:string) 
                        (src:MarkdownDoc) : DocMonad<PdfDoc, #WordDocument.IWordHandle> =
        docMonad { 
            let docName = Path.ChangeExtension(outputPdfName, "docx")
            let! doc = markdownToWordAs docName src
            return! WordDocument.exportPdfAs outputPdfName doc
         }

    /// This uses MS Word as to render in intermediate docx file to Pdf.
    let markdownToPdf (source : MarkdownDoc) : DocMonad<PdfDoc, #WordDocument.IWordHandle> =
        docMonad {
            let! sourceName = getDocumentFileName source
            let fileName = Path.ChangeExtension(sourceName, "pdf")
            return! markdownToPdfAs fileName source
        }



    type ContentsConfig = Contents.ContentsConfig

    /// Make a title page PDF.
    /// Render to docx then use Word to render to PDF.
    let makeTableOfContents (config:ContentsConfig) 
                            (col:PdfCollection) : DocMonad<PdfDoc, #WordDocument.IWordHandle> =
        Contents.genTableOfContents markdownToPdf config col 

    type PhotoBookConfig = PhotoBook.PhotoBookConfig

    let makePhotoBook (config:PhotoBookConfig) : DocMonad<PdfDoc, 'res> =
        PhotoBook.genPhotoBook markdownToPdf config

    /// Prefix the Pdf with a title page.
    /// Render to docx then use Word to render to PDF.
    let prefixWithTitlePage (title:string) 
                                    (body: Markdown option) 
                                    (pdf:PdfDoc) : DocMonad<PdfDoc, #WordDocument.IWordHandle> =
        TitlePage.genPrefixWithTitlePage markdownToPdf title body pdf

