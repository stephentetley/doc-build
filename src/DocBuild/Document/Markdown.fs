// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Document


module Markdown = 

    open System.IO
    open MarkdownDoc

    open DocBuild.Base
    open DocBuild.Base.Internal


    // ************************************************************************
    // Retrieve Custom styles
    
    let getCustomStyles () : DocMonad<WordDoc option, 'userRes> = 
        docMonad { 
            match! asks(fun env -> env.PandocOpts.CustomStylesDocx)  with
            | None -> return None
            | Some path -> 
                if isAbsolutePath path then 
                    return! getWordDoc path |>> Some
                else
                    return! getIncludeWordDoc path |>> Some
        }
                    
    let private getCustomStylesPath () : DocMonad<string option, 'userRes> = 
        getCustomStyles () 
            |>> Option.map (fun (doc:WordDoc) -> doc.AbsolutePath)



    // ************************************************************************
    // Save output from MarkdownDoc

    /// Output a Markdown doc to file.
    let saveMarkdown (outputRelName:string) 
                     (markdown:Markdown) : DocMonad<MarkdownDoc, 'userRes> = 
        docMonad { 
            let! outputAbsPath = extendWorkingPath outputRelName
            let _ = markdown.Save outputAbsPath 
            return! getMarkdownDoc outputAbsPath
        }

    // ************************************************************************
    // Export

    /// Requires pandoc
    let markdownToWordAs (outputRelName:string) 
                         (src:MarkdownDoc) : DocMonad<WordDoc, 'userRes> =
        docMonad { 
            let! outputAbsPath = extendWorkingPath outputRelName
            let! styles = getCustomStylesPath () 
            let command = 
                PandocPrim.outputDocxCommand styles [] src.AbsolutePath outputAbsPath
            let! _ = execPandoc command
            return! getWordDoc outputAbsPath
         }

    /// Requires pandoc
    let markdownToWord (src:MarkdownDoc) : DocMonad<WordDoc, 'userRes> =
        let outputName = Path.ChangeExtension(src.AbsolutePath, "docx") |> Path.GetFileName
        markdownToWordAs outputName src




    // ************************************************************************
    // Find and replace

    let findReplaceAs (searches:SearchList) 
                      (outputRelName:string) 
                      (src:MarkdownDoc) : DocMonad<MarkdownDoc, 'userRes> = 
        docMonad { 
            let! outputAbsPath = extendWorkingPath outputRelName
            let original = File.ReadAllText(src.AbsolutePath)
            let action (source:string) (searchText:string, replaceText:string) = 
               source.Replace(searchText, replaceText)
            let final = List.fold action original searches
            let _ = File.WriteAllText(outputAbsPath, final)
            return! getMarkdownDoc outputAbsPath
        }


    let findReplace (searches:SearchList)
                    (src:MarkdownDoc) : DocMonad<MarkdownDoc, 'userRes> = 
        findReplaceAs searches src.FileName src
