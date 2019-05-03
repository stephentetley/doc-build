// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Office

[<RequireQualifiedAccess>]
module PandocWordShim = 

    open System.IO

    // Open Office at .Interop rather than .Word then the Word API 
    // has to be qualified.
    open Microsoft.Office.Interop

    open MarkdownDoc

    open DocBuild.Base
    open DocBuild.Base.DocMonad
    open DocBuild.Document.Markdown
    open DocBuild.Extra
    
    open DocBuild.Office

    let asksCustomStyles () : DocMonad<'res, WordDoc option> = 
        docMonad {
            match! asks (fun env -> env.PandocOpts.CustomStylesDocx) with 
            | None -> return None 
            | Some fileName -> return! getIncludeWordDoc fileName |>> Some
        }


    // ************************************************************************
    // Export to Pdf with Pandoc and MS Word

    /// This uses MS Word as to render in intermediate docx file to Pdf.
    let markdownToPdfAs (outputPdfName:string) 
                        (src:MarkdownDoc) : DocMonad<#WordDocument.IWordHandle,PdfDoc> =
        docMonad { 
            let docName = Path.ChangeExtension(outputPdfName, "docx")
            let! doc = markdownToWordAs docName src
            return! WordDocument.exportPdfAs outputPdfName doc
         }

    /// This uses MS Word as to render in intermediate docx file to Pdf.
    let markdownToPdf (src:MarkdownDoc) : DocMonad<#WordDocument.IWordHandle,PdfDoc> =
        let fileName = Path.ChangeExtension(src.FileName, "pdf")
        markdownToPdfAs fileName src



    type ContentsConfig = Contents.ContentsConfig

    /// Make a title page PDF.
    /// Render to docx then use Word to render to PDF.
    let makeTableOfContents (config:ContentsConfig) 
                            (col:PdfCollection) : DocMonad<#WordDocument.IWordHandle,PdfDoc> =
        Contents.genTableOfContents markdownToPdf config col 

    type PhotoBookConfig = PhotoBook.PhotoBookConfig

    let makePhotoBook (config:PhotoBookConfig) : DocMonad<'res, PdfDoc> =
        PhotoBook.genPhotoBook markdownToPdf config

    /// Prefix the Pdf with a title page.
    /// Render to docx then use Word to render to PDF.
    let prefixWithTitlePage (title:string) 
                                    (body: Markdown option) 
                                    (pdf:PdfDoc) : DocMonad<#WordDocument.IWordHandle,PdfDoc> =
        TitlePage.genPrefixWithTitlePage markdownToPdf title body pdf

