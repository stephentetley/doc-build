// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocBuild.Extra


module PandocTeXShim = 

    open System.IO

    open SLFormat.CommandOptions        // lib: sl-format
    open MarkdownDoc                    // lib: MarkdownDoc

    open DocBuild.Base
    open DocBuild.Base.Internal
    open DocBuild.Extra


    // ************************************************************************
    // Export to Pdf with Pandoc (and TeX)


    /// The specific TeX backend is set in DocBuildEnv, generally you 
    /// should use "pdflatex".
    let markdownToPdfAs (outputRelName:string) 
                        (src:MarkdownDoc) : DocMonad<PdfDoc, 'userRes> =
        docMonad { 
            let! outputAbsPath = extendWorkingPath outputRelName
            let! pdfEngine = asks (fun env -> env.PandocOpts.PdfEngine)       
            let command = 
                PandocPrim.outputPdfCommand pdfEngine [] src.AbsolutePath outputAbsPath
            
            let! _ = execPandoc command
            return! getPdfDoc outputAbsPath
         }


    let markdownToPdf (src:MarkdownDoc) : DocMonad<PdfDoc, 'userRes> =
        let outputName = Path.ChangeExtension(src.AbsolutePath, "pdf") |> Path.GetFileName
        markdownToPdfAs outputName src


    type ContentsConfig = Contents.ContentsConfig

    /// Make a title page PDF.
    /// Use Pandoc to render to PDF via TeX.
    /// TeX must be installed and callable by Pandoc.
    let makeTableOfContents (config:ContentsConfig) 
                            (col:PdfCollection) : DocMonad<PdfDoc, 'userRes> =
        Contents.genTableOfContents markdownToPdf config col


    type PhotoBookConfig = PhotoBook.PhotoBookConfig


    /// Make a 'photo book'.
    /// Use Pandoc to render to PDF via TeX.
    /// TeX must be installed and callable by Pandoc.
    let makePhotoBook (config:PhotoBookConfig) : DocMonad<PdfDoc, 'userRes> =
        PhotoBook.genPhotoBook markdownToPdf config

    /// Prefix the Pdf with a title page.
    /// Use Pandoc to render to PDF via TeX.
    /// TeX must be installed and callable by Pandoc.
    let prefixWithTitlePage (title:string) 
                            (body: Markdown option) 
                            (pdf:PdfDoc) : DocMonad<PdfDoc, 'userRes> =
        TitlePage.genPrefixWithTitlePage markdownToPdf title body pdf
