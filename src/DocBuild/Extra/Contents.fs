// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocBuild.Extra

module Contents = 


    open MarkdownDoc
    open MarkdownDoc.Pandoc

    open DocBuild.Base
    open DocBuild.Base.DocMonad
    open DocBuild.Base.DocMonadOperators
    open DocBuild.Base.Collection


    open DocBuild.Document

    // TODO
    let makeContents (pdfs:PdfCollection): DocMonad<'res,MarkdownDoc> =
        throwError "TODO"
