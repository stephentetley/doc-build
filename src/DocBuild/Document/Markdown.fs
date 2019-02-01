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
    let saveMarkdown (outputAbsPath:string) 
                     (markdown:Markdown) : DocMonad<'res,MarkdownFile> = 
        docMonad { 
            do! assertIsWorkingPath outputAbsPath
            let _ = markdown.Save outputAbsPath 
            return! workingMarkdownFile outputAbsPath
        }

    // ************************************************************************
    // Export

    let markdownToWordAs (customStyles:WordFile option)
                         (outputAbsPath:string) 
                         (src:MarkdownFile) : DocMonad<'res,WordFile> =
        docMonad { 
            do! assertIsWorkingPath outputAbsPath
            let styles = customStyles |> Option.map (fun doc -> doc.LocalPath) 
            let command = 
                PandocPrim.outputDocxCommand styles src.LocalPath outputAbsPath
            let! _ = execPandoc command
            return! workingWordFile outputAbsPath
         }


    let markdownToWord (customStyles:WordFile option) 
                       (src:MarkdownFile) : DocMonad<'res,WordFile> =
        let outputFile = Path.ChangeExtension(src.LocalPath, "docx")
        markdownToWordAs customStyles outputFile src



    // ************************************************************************
    // Find and replace

    let findReplaceAs (searches:SearchList) 
                      (outputAbsPath:string) 
                      (src:MarkdownFile) : DocMonad<'res,MarkdownFile> = 
        docMonad { 
            do! assertIsWorkingPath outputAbsPath
            let original = File.ReadAllText(src.LocalPath)
            let action (source:string) (searchText:string, replaceText:string) = 
               source.Replace(searchText, replaceText)
            let final = List.fold action original searches
            let _ = File.WriteAllText(outputAbsPath, final)
            return! workingMarkdownFile outputAbsPath
        }


    let findReplace (searches:SearchList)
                    (src:MarkdownFile) : DocMonad<'res,MarkdownFile> = 
        findReplaceAs searches src.LocalPath src
