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
    let saveMarkdown (outputName:string) 
                     (markdown:Markdown) : DocMonad<'res,MarkdownFile> = 
        docMonad { 
            let! outputPath = getOutputPath outputName
            let _ = markdown.Save outputPath 
            let! md = workingMarkdownFile outputPath
            return md
        }

    // ************************************************************************
    // Export

    let markdownToWordAs (customStyles:WordFile option)
                         (outputName:string) 
                         (src:MarkdownFile) : DocMonad<'res,WordFile> =
        docMonad { 
            let! outputPath = getOutputPath outputName
            let styles = customStyles |> Option.map (fun doc -> doc.AbsolutePath) 
            let command = 
                PandocPrim.outputDocxCommand styles src.AbsolutePath outputPath
            let! _ = execPandoc command
            let! docx = workingWordFile outputName
            return docx
         }


    let markdownToWord (customStyles:WordFile option) 
                       (src:MarkdownFile) : DocMonad<'res,WordFile> =
        let outputFile = Path.ChangeExtension(src.AbsolutePath, "docx")
        markdownToWordAs customStyles outputFile src



    // ************************************************************************
    // Find and replace

    let findReplaceAs (searches:SearchList) 
                      (outputName:string) 
                      (src:MarkdownFile) : DocMonad<'res,MarkdownFile> = 
        docMonad { 
            let! outputPath = getOutputPath outputName
            let original = File.ReadAllText(src.AbsolutePath)
            let action (source:string) (searchText:string, replaceText:string) = 
               source.Replace(searchText, replaceText)
            let final = List.fold action original searches
            let _ = File.WriteAllText(outputPath, final)
            let! md = workingMarkdownFile outputName
            return md
        }


    let findReplace (searches:SearchList)
                    (src:MarkdownFile) : DocMonad<'res,MarkdownFile> = 
        findReplaceAs searches src.FileName src
