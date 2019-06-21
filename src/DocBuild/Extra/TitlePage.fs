// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocBuild.Extra

module TitlePage = 

    open System.IO

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


    let makeTitlePage (config:TitlePageConfig) : DocMonad<MarkdownDoc, 'userRes> =
        docMonad {
            let mdDoc = genMarkdown config.Title config.DocBody
            return! Markdown.saveMarkdown config.RelativeOutputName mdDoc
        }

    let genPrefixWithTitlePage (render: MarkdownDoc -> DocMonad<PdfDoc, 'userRes>)
                               (title:string) 
                               (body: Markdown option) 
                               (source : PdfDoc) : DocMonad<PdfDoc, 'userRes> =
        docMonad {
            // TODO this is imperminent, need an easy genfile function
            let temp = "title.temp.md"    
            let! md = makeTitlePage { Title = title; DocBody = body; RelativeOutputName = temp }
            let! title = render md
            let! sourceName = getDocumentFileName source
            let outName = modifyFileName (fun s -> s + "+title") sourceName 
            return! pdftkConcatPdfs outName (Collection.ofList [title; source]) |>> setTitle source.Title
        }



        
