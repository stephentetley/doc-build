// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocBuild.Extra

module PandocTeXShim = 

    open MarkdownDoc            /// lib: MarkdownDoc

    open DocBuild.Base
    open DocBuild.Base.DocMonad
    open DocBuild.Document.Markdown
    open DocBuild.Extra

    type ContentsConfig = Contents.ContentsConfig

    /// Make a title page PDF.
    /// Use Pandoc to render to PDF via TeX.
    /// TeX must be installed and callable by Pandoc.
    let makeTableOfContents (config:ContentsConfig) 
                            (col:PdfCollection) : DocMonad<'userRes, PdfDoc> =
        Contents.genTableOfContents markdownToTeXToPdf config col


    type PhotoBookConfig = PhotoBook.PhotoBookConfig


    /// Make a 'photo book'.
    /// Use Pandoc to render to PDF via TeX.
    /// TeX must be installed and callable by Pandoc.
    let makePhotoBook (config:PhotoBookConfig) : DocMonad<'userRes, PdfDoc option> =
        PhotoBook.genPhotoBook markdownToTeXToPdf config

    /// Prefix the Pdf with a title page.
    /// Use Pandoc to render to PDF via TeX.
    /// TeX must be installed and callable by Pandoc.
    let prefixWithTitlePage (title:string) 
                            (body: Markdown option) 
                            (pdf:PdfDoc) : DocMonad<'userRes, PdfDoc> =
        TitlePage.genPrefixWithTitlePage markdownToTeXToPdf title body pdf
