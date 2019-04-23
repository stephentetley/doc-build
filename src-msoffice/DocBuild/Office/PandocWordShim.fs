// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Office


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
        asks (fun env -> env.PandocOpts.CustomStylesDocx) >>= fun opt -> 
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
        let outputFile = Path.ChangeExtension(src.AbsolutePath, "pdf")
        markdownToWordToPdfAs outputFile src



    type ContentsConfig = Contents.ContentsConfig

    /// Make a title page PDF.
    /// Render to docx then use Word to render to PDF.
    let makeTableOfContents (config:ContentsConfig) 
                            (col:PdfCollection) : DocMonad<#WordDocument.IWordHandle,PdfDoc> =
        Contents.genTableOfContents markdownToWordToPdf config col 

    type PhotoBookConfig = PhotoBook.PhotoBookConfig

    let makePhotoBook (config:PhotoBookConfig) : DocMonad<'res, PdfDoc option> =
        PhotoBook.genPhotoBook markdownToWordToPdf config

    /// Prefix the Pdf with a title page.
    /// Render to docx then use Word to render to PDF.
    let prefixWithTitlePage (title:string) 
                                    (body: Markdown option) 
                                    (pdf:PdfDoc) : DocMonad<#WordDocument.IWordHandle,PdfDoc> =
        TitlePage.genPrefixWithTitlePage markdownToWordToPdf title body pdf

