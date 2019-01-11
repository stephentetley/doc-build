// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Document


module Markdown = 

    open System.IO

    open DocBuild.Base
    open DocBuild.Base.DocMonad
    open DocBuild.Raw
    open System

    // ************************************************************************
    // Export

    let markdownToWordAs (src:MarkdownFile) (outputFile:string) : DocMonad<WordFile> =
        docMonad { 
            let! styles = asks (fun env -> env.PandocReferenceDoc)
            let command = 
                PandocPrim.outputDocxCommand styles  src.Path outputFile
            let! _ = execPandoc command
            let! docx = getWordFile outputFile
            return docx
         }


    let markdownToWord (src:MarkdownFile) : DocMonad<WordFile> =
        let outputFile = Path.ChangeExtension(src.Path, "docx")
        markdownToWordAs src outputFile



    // ************************************************************************
    // Find and replace

    let findReplaceAs (src:MarkdownFile) (searches:SearchList) (outputFile:string) : DocMonad<MarkdownFile> = 
        let original = File.ReadAllText(src.Path)
        let action (source:string) (searchText:string, replaceText:string) = 
           source.Replace(searchText, replaceText)
        let final = List.fold action original searches
        File.WriteAllText(outputFile, final)
        getMarkdownFile outputFile


    let findReplace (src:MarkdownFile) (searches:SearchList)  : DocMonad<MarkdownFile> = 
        findReplaceAs src searches src.NextTempName 
