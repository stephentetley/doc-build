// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Document


module Markdown = 

    open System.IO
    open MarkdownDoc.Markdown

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
            |>> Option.bind (fun (doc:WordDoc) -> doc.AbsolutePath)



    // ************************************************************************
    // Save output from MarkdownDoc

    /// Output a Markdown doc to file.
    let saveMarkdown (outputRelName:string) 
                     (markdown:Markdown) : DocMonad<MarkdownDoc, 'userRes> = 
        docMonad { 
            let! outputAbsPath = extendWorkingPath outputRelName
            let _ = writeMarkdown 180 markdown outputAbsPath 
            return! getMarkdownDoc outputAbsPath
        }

    // ************************************************************************
    // Export

    /// Requires pandoc
    let markdownToWordAs (outputRelName : string) 
                         (source : MarkdownDoc) : DocMonad<WordDoc, 'userRes> =
        docMonad { 
            let! outputAbsPath = extendWorkingPath outputRelName
            let! sourcePath = getDocumentPath source
            let! styles = getCustomStylesPath () 
            let command = 
                PandocPrim.outputDocxCommand styles [] sourcePath outputAbsPath
            let! _ = execPandoc command
            return! getWordDoc outputAbsPath
         }

    /// Requires pandoc
    let markdownToWord (source : MarkdownDoc) : DocMonad<WordDoc, 'userRes> =
        docMonad { 
            let! sourceName = getDocumentFileName source
            let outputName = Path.ChangeExtension(sourceName, "docx") |> Path.GetFileName
            return! markdownToWordAs outputName source
        }
        
        
        




    // ************************************************************************
    // Find and replace

    let findReplaceAs (searches : SearchList) 
                      (outputRelName : string) 
                      (source : MarkdownDoc) : DocMonad<MarkdownDoc, 'userRes> = 
        docMonad { 
            let! outputAbsPath = extendWorkingPath outputRelName
            let! sourcePath = getDocumentPath source
            let original = File.ReadAllText(sourcePath)
            let action (source:string) (searchText:string, replaceText:string) = 
               source.Replace(searchText, replaceText)
            let final = List.fold action original searches
            let _ = File.WriteAllText(outputAbsPath, final)
            return! getMarkdownDoc outputAbsPath
        }


    let findReplace (searches : SearchList)
                    (source : MarkdownDoc) : DocMonad<MarkdownDoc, 'userRes> = 
        docMonad { 
            let! sourceName = getDocumentFileName source
            return! findReplaceAs searches sourceName source
        }
