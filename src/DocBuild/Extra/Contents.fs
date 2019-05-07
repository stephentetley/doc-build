// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocBuild.Extra

module Contents = 

    open MarkdownDoc
    
    open DocBuild.Base

    open DocBuild.Document
    open DocBuild.Document.Pdf

    // We can render a preliminary version of contents to get its length,
    // or choose to render a certain number of items per page and know 
    // the length in advance (caveat, expect one entry = one line)

    type private DocInfo = 
        { Title: string 
          PageCount: int }

    let private contentsTable (start:int) (infos:DocInfo list) : Markdown = 
        List.fold (fun (i,ac) (info:DocInfo) -> 
                        let d1 = ac ^@^ h2 (text info.Title ^+^ text "..." ^+^ formatted "%i" i)
                        (i + info.PageCount, d1))
                  (start, Markdown.empty)
                  infos
            |> snd

    let private genMarkdown (prologLength:int) (infos:DocInfo list) : Markdown = 
        h1 (text "Contents") ^@^ contentsTable (prologLength+1) infos 


    /// Prolog size is the number of pages in any coversheet 
    /// etc. before the table-of-contents.
    type ContentsConfig = 
        { PrologLength: int
          RelativeOutputName: string }

    let private getInfo (pdf:PdfDoc) : DocMonad<DocInfo, 'userRes> =
        docMonad { 
            let! count = countPages pdf
            return { Title = pdf.Title; PageCount = count}        
        }

    // TODO - render a dummy doc, to get length of contents

    let makeContents1 (config:ContentsConfig) 
                      (col:PdfCollection) : DocMonad<MarkdownDoc, 'userRes> =
        docMonad {
            let! (infos:DocInfo list) = mapM getInfo col.Elements
            let mdDoc = genMarkdown config.PrologLength infos
            return! Markdown.saveMarkdown config.RelativeOutputName mdDoc           
        }

    let genTableOfContents (render: MarkdownDoc -> DocMonad<PdfDoc, 'userRes>)
                           (config:ContentsConfig) 
                           (col:PdfCollection) : DocMonad<PdfDoc, 'userRes> =
        docMonad {
            let config1 = { config with RelativeOutputName = "contents-zero.md" }
            let! tocTemp = makeContents1 config1 col >>= render
            let! newPrologLength = countPages tocTemp |>> (fun i -> i + config.PrologLength)
            let config2 =  { config with PrologLength = newPrologLength }
            let! tocFinal = makeContents1 config2 col
            return! render tocFinal
        }

