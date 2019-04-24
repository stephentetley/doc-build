// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Document


module Markdown = 

    open System.IO
    open SLFormat.CommandOptions
    open MarkdownDoc

    open DocBuild.Base
    open DocBuild.Base.Internal


    // ************************************************************************
    // Retrieve Custom styles
    
    let getCustomStyles () : DocMonad<'userRes, WordDoc option> = 
        docMonad { 
            match! asks(fun env -> env.PandocOpts.CustomStylesDocx)  with
            | None -> return None
            | Some path -> 
                if isAbsolutePath path then 
                    return! getWordDoc path |>> Some
                else
                    return! getIncludeWordDoc path |>> Some
        }
                    
    let private getCustomStylesPath () : DocMonad<'userRes, string option> = 
        getCustomStyles () 
            |>> Option.map (fun (doc:WordDoc) -> doc.AbsolutePath)



    // ************************************************************************
    // Save output from MarkdownDoc

    /// Output a Markdown doc to file.
    let saveMarkdown (outputAbsPath:string) 
                     (markdown:Markdown) : DocMonad<'userRes,MarkdownDoc> = 
        docMonad { 
            do! assertIsWorkingPath outputAbsPath
            let _ = markdown.Save outputAbsPath 
            return! getWorkingMarkdownDoc outputAbsPath
        }

    // ************************************************************************
    // Export

    /// Requires pandoc
    let markdownToWordAs (outputAbsPath:string) 
                         (src:MarkdownDoc) : DocMonad<'userRes,WordDoc> =
        docMonad { 
            do! assertIsWorkingPath outputAbsPath
            let! styles = getCustomStylesPath () 
            let command = 
                PandocPrim.outputDocxCommand styles [] src.AbsolutePath outputAbsPath
            let! _ = execPandoc command
            return! getWorkingWordDoc outputAbsPath
         }

    /// Requires pandoc
    let markdownToWord (src:MarkdownDoc) : DocMonad<'userRes,WordDoc> =
        let outputFile = Path.ChangeExtension(src.AbsolutePath, "docx")
        markdownToWordAs outputFile src


    // ************************************************************************
    // Export to Pdf with Pandoc (and TeX)


    ///  Specific TeX backend is set in DocBuildEnv, generally you 
    /// should use "pdflatex".
    let markdownToTeXToPdfAs (outputAbsPath:string) 
                             (src:MarkdownDoc) : DocMonad<'userRes,PdfDoc> =
        docMonad { 
            do! assertIsWorkingPath outputAbsPath
            let! pdfEngine = asks (fun env -> env.PandocOpts.PdfEngine)       
            let command = 
                PandocPrim.outputPdfCommand pdfEngine [] src.AbsolutePath outputAbsPath
            printfn "// %s" (arguments command)
            let! _ = execPandoc command
            return! getWorkingPdfDoc outputAbsPath
         }


    let markdownToTeXToPdf (src:MarkdownDoc) : DocMonad<'userRes,PdfDoc> =
        let outputFile = Path.ChangeExtension(src.AbsolutePath, "pdf")
        markdownToTeXToPdfAs outputFile src

    // ************************************************************************
    // Find and replace

    let findReplaceAs (searches:SearchList) 
                      (outputAbsPath:string) 
                      (src:MarkdownDoc) : DocMonad<'userRes,MarkdownDoc> = 
        docMonad { 
            do! assertIsWorkingPath outputAbsPath
            let original = File.ReadAllText(src.AbsolutePath)
            let action (source:string) (searchText:string, replaceText:string) = 
               source.Replace(searchText, replaceText)
            let final = List.fold action original searches
            let _ = File.WriteAllText(outputAbsPath, final)
            return! getWorkingMarkdownDoc outputAbsPath
        }


    let findReplace (searches:SearchList)
                    (src:MarkdownDoc) : DocMonad<'userRes,MarkdownDoc> = 
        findReplaceAs searches src.AbsolutePath src
