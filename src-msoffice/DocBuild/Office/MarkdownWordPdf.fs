﻿// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Office


module MarkdownWordPdf = 

    open System.IO

    // Open at .Interop rather than .Word then the Word API has to be qualified
    open Microsoft.Office.Interop

    open MarkdownDoc

    open DocBuild.Base
    open DocBuild.Base.DocMonad
    open DocBuild.Base.DocMonadOperators
    open DocBuild.Document.Markdown
    open DocBuild.Extra.Contents
    open DocBuild.Extra.TitlePage
    open DocBuild.Extra.PhotoBook
    
    open DocBuild.Office

    let asksCustomStyles () : DocMonad<'res, WordDoc option> = 
        asks (fun env -> env.CustomStylesDocx) >>= fun opt -> 
        match opt with 
        | None -> mreturn None 
        | Some absPath -> getWordDoc absPath |>> Some


    // ************************************************************************
    // Export to Pdf with Pandoc (and TeX)


    let markdownToWordToPdfAs (outputAbsPath:string) 
                              (src:MarkdownDoc) : DocMonad<#WordDocument.IWordHandle,PdfDoc> =
        docMonad { 
            let docAbsPath = Path.ChangeExtension(outputAbsPath, "docx")
            let! doc = markdownToWordAs docAbsPath src
            return! WordDocument.exportPdfAs outputAbsPath doc
         }


    let markdownToWordToPdf (src:MarkdownDoc)  : DocMonad<#WordDocument.IWordHandle,PdfDoc> =
        let outputFile = Path.ChangeExtension(src.LocalPath, "pdf")
        markdownToWordToPdfAs outputFile src


    /// Prefix the Pdf with a title page.
    /// Render to docx then use Word to render to PDF.
    let prefixWithTitlePageWithWord (title:string) 
                                    (body: Markdown option) 
                                    (pdf:PdfDoc) : DocMonad<#WordDocument.IWordHandle,PdfDoc> =
        genPrefixWithTitlePage markdownToWordToPdf title body pdf


    /// Make a title page PDF.
    /// Render to docx then use Word to render to PDF.
    let makeContentsWithWord (config:ContentsConfig) 
                             (col:PdfCollection) : DocMonad<#WordDocument.IWordHandle,PdfDoc> =
        genMakeContents markdownToWordToPdf config col 


    let makePhotoBookWithWord (config:PhotoBookConfig) : DocMonad<'res, PdfDoc option> =
        genPhotoBook markdownToWordToPdf config