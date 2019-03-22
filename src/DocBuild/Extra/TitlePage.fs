// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocBuild.Extra

module TitlePage = 

    open MarkdownDoc
    
    open DocBuild.Base
    open DocBuild.Base.DocMonad

    open DocBuild.Document
    open DocBuild.Document.Pdf
    open DocBuild.Document.Markdown


    let private genMarkdown (title:string) 
                            (body:Markdown option) : Markdown = 
        let d1 = h1 (text title)
        match body with
        | None -> d1
        | Some d2 -> d1 ^@^ d2




    type TitlePageConfig = 
        { Title: string
          DocBody: Markdown option
          RelativeOutputName: string }


    let makeTitlePage (config:TitlePageConfig) : DocMonad<'res, MarkdownDoc> =
        docMonad {
            let mdDoc = genMarkdown config.Title config.DocBody
            let! outputAbsPath = extendWorkingPath config.RelativeOutputName
            let! _ = Markdown.saveMarkdown outputAbsPath mdDoc
            return! workingMarkdownDoc outputAbsPath
        }

    let genPrefixWithTitlePage (render: MarkdownDoc -> DocMonad<'res,PdfDoc>)
                               (title:string) 
                               (body: Markdown option) 
                               (pdf:PdfDoc) : DocMonad<'res,PdfDoc> =
        docMonad {
            // TODO this is imperminent, need an easy genfile function
            let temp = "title.temp.md"    
            let! md = makeTitlePage { Title = title; DocBody = body; RelativeOutputName = temp }
            let! title = render md
            let outPath = modifyFileName (fun s -> s + "+title") pdf.LocalPath 
            let! ans = pdftkConcatPdfs (Collection.fromList [title; pdf]) outPath
            return ans |> setTitle pdf.Title
        }


    let prefixWithTitlePageWithTeX (title:string) 
                                   (body: Markdown option) 
                                   (pdf:PdfDoc) : DocMonad<'res,PdfDoc> =
        genPrefixWithTitlePage markdownToTeXToPdf title body pdf
        
