// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Document


module Markdown = 

    open System.IO

    open DocBuild.Base
    open DocBuild.Base.DocMonad
    open DocBuild.Raw



    let markdownToWordAs (src:MarkdownFile) (outputFile:string) : DocMonad<WordFile> =
        docMonad { 
            let! styles = asks (fun env -> env.PandocReferenceDoc)
            let command = 
                PandocPrim.outputDocxCommand styles  src.Path outputFile
            let! _ = execPandoc command
            let! docx = wordFile outputFile
            return docx
         }


    let markdownToWord (src:MarkdownFile) : DocMonad<WordFile> =
        let outputFile = Path.ChangeExtension(src.Path, "docx")
        markdownToWordAs src outputFile

