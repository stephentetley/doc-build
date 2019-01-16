// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Document


module Markdown = 

    open System.IO

    open MarkdownDoc

    open DocBuild.Base
    open DocBuild.Base.DocMonad
    open DocBuild.Raw
    // open System


    // ************************************************************************
    // Save output from MarkdownDoc

    /// Output a Markdown doc to file.
    let saveMarkdown (markdown:Markdown) 
                     (outputFile:string) : DocMonad<'res,MarkdownFile> = 
        markdown.Save outputFile 
        getMarkdownFile outputFile

    // ************************************************************************
    // Export

    let markdownToWordAs (src:MarkdownFile) 
                         (outputFile:string) : DocMonad<'res,WordFile> =
        docMonad { 
            let! styles = asks (fun env -> env.PandocReferenceDoc)
            let command = 
                PandocPrim.outputDocxCommand styles  src.Path outputFile
            let! _ = execPandoc command
            let! docx = getWordFile outputFile
            return docx
         }


    let markdownToWord (src:MarkdownFile) : DocMonad<'res,WordFile> =
        let outputFile = Path.ChangeExtension(src.Path, "docx")
        markdownToWordAs src outputFile



    // ************************************************************************
    // Find and replace

    let findReplaceAs (src:MarkdownFile) 
                      (searches:SearchList) 
                      (outputFile:string) : DocMonad<'res,MarkdownFile> = 
        let original = File.ReadAllText(src.Path)
        let action (source:string) (searchText:string, replaceText:string) = 
           source.Replace(searchText, replaceText)
        let final = List.fold action original searches
        File.WriteAllText(outputFile, final)
        getMarkdownFile outputFile


    let findReplace (src:MarkdownFile) 
                    (searches:SearchList)  : DocMonad<'res,MarkdownFile> = 
        findReplaceAs src searches src.NextTempName 
