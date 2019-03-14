// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocBuild.Extra

module TitlePage = 

    open MarkdownDoc
    
    open DocBuild.Base
    open DocBuild.Base.DocMonad

    open DocBuild.Document
    open DocBuild.Document.Pdf

    // We can render a preliminary version of contents to get its length.

    type private DocInfo = 
        { Title: string 
          PageCount: int }


    let private genMarkdown (title:string) : Markdown = 
        h1 (text title)



    type TitlePageConfig = 
        { Title: string
          RelativeOutputName: string }


    let makeTitlePage (config:TitlePageConfig) : DocMonad<'res, MarkdownDoc> =
        docMonad {
            let mdDoc = genMarkdown config.Title
            let! outputAbsPath = extendWorkingPath config.RelativeOutputName
            let! _ = Markdown.saveMarkdown outputAbsPath mdDoc
            return! workingMarkdownDoc outputAbsPath
        }