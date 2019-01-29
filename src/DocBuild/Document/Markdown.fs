// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Document


module Markdown = 

    open System.IO

    open MarkdownDoc

    open DocBuild.Base
    open DocBuild.Base.DocMonad
    open DocBuild.Raw


    // ************************************************************************
    // Save output from MarkdownDoc

    /// Output a Markdown doc to file.
    let saveMarkdown (outputFile:string) 
                     (markdown:Markdown) : DocMonad<'res,MarkdownFile> = 
        markdown.Save outputFile 
        getMarkdownFile outputFile

    // ************************************************************************
    // Export

    let markdownToWordAs (customStyles:WordFile option)
                         (outputFile:string) 
                         (src:MarkdownFile) : DocMonad<'res,WordFile> =
        docMonad { 
            let styles = customStyles |> Option.map (fun doc -> doc.Path.AbsolutePath) 
            let command = 
                PandocPrim.outputDocxCommand styles src.Path.AbsolutePath outputFile
            let! _ = execPandoc command
            let! docx = getWordFile outputFile
            return docx
         }


    let markdownToWord (customStyles:WordFile option) 
                       (src:MarkdownFile) : DocMonad<'res,WordFile> =
        let outputFile = Path.ChangeExtension(src.Path.AbsolutePath, "docx")
        markdownToWordAs customStyles outputFile src



    // ************************************************************************
    // Find and replace

    let findReplaceAs (searches:SearchList) 
                      (outputFile:string) 
                      (src:MarkdownFile) : DocMonad<'res,MarkdownFile> = 
        let original = File.ReadAllText(src.Path.AbsolutePath)
        let action (source:string) (searchText:string, replaceText:string) = 
           source.Replace(searchText, replaceText)
        let final = List.fold action original searches
        File.WriteAllText(outputFile, final)
        getMarkdownFile outputFile


    let findReplace (searches:SearchList)
                    (src:MarkdownFile) : DocMonad<'res,MarkdownFile> = 
        findReplaceAs searches src.NextTempName.AbsolutePath src
