// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Document


module Markdown = 

    open System.IO
    open SLFormat.CommandOptions
    open MarkdownDoc

    open DocBuild.Base
    open DocBuild.Base.DocMonad
    open DocBuild.Raw

    // ************************************************************************
    // Retrieve Custom styles
    
    let getCustomStyles () : DocMonad<'res, WordDoc option> = 
        docMonad { 
            match! asks(fun env -> env.PandocOpts.CustomStylesDocx)  with
            | None -> return None
            | Some path -> 
                if isAbsolutePath path then 
                    return! getWordDoc path |>> Some
                else
                    return! includeWordDoc path |>> Some
        }
                    
    let private getCustomStylesPath () : DocMonad<'res, string option> = 
        getCustomStyles () 
            |>> Option.map (fun (doc:WordDoc) -> doc.LocalPath)



    // ************************************************************************
    // Save output from MarkdownDoc

    /// Output a Markdown doc to file.
    let saveMarkdown (outputAbsPath:string) 
                     (markdown:Markdown) : DocMonad<'res,MarkdownDoc> = 
        docMonad { 
            do! assertIsWorkingPath outputAbsPath
            let _ = markdown.Save outputAbsPath 
            return! workingMarkdownDoc outputAbsPath
        }

    // ************************************************************************
    // Export

    /// Requires pandoc
    let markdownToWordAs (outputAbsPath:string) 
                         (src:MarkdownDoc) : DocMonad<'res,WordDoc> =
        docMonad { 
            do! assertIsWorkingPath outputAbsPath
            let! styles = getCustomStylesPath () 
            let command = 
                PandocPrim.outputDocxCommand styles [] src.LocalPath outputAbsPath
            let! _ = execPandoc command
            return! workingWordDoc outputAbsPath
         }

    /// Requires pandoc
    let markdownToWord (src:MarkdownDoc) : DocMonad<'res,WordDoc> =
        let outputFile = Path.ChangeExtension(src.LocalPath, "docx")
        markdownToWordAs outputFile src


    // ************************************************************************
    // Export to Pdf with Pandoc (and TeX)


    ///  Specific TeX backend is set in DocBuildEnv, generally you 
    /// should use "pdflatex".
    let markdownToTeXToPdfAs (outputAbsPath:string) 
                             (src:MarkdownDoc) : DocMonad<'res,PdfDoc> =
        docMonad { 
            do! assertIsWorkingPath outputAbsPath
            let! pdfEngine = asks (fun env -> env.PandocOpts.PdfEngine)       
            let command = 
                PandocPrim.outputPdfCommand pdfEngine [] src.LocalPath outputAbsPath
            printfn "// %s" (arguments command)
            let! _ = execPandoc command
            return! workingPdfDoc outputAbsPath
         }


    let markdownToTeXToPdf (src:MarkdownDoc) : DocMonad<'res,PdfDoc> =
        let outputFile = Path.ChangeExtension(src.LocalPath, "pdf")
        markdownToTeXToPdfAs outputFile src

    // ************************************************************************
    // Find and replace

    let findReplaceAs (searches:SearchList) 
                      (outputAbsPath:string) 
                      (src:MarkdownDoc) : DocMonad<'res,MarkdownDoc> = 
        docMonad { 
            do! assertIsWorkingPath outputAbsPath
            let original = File.ReadAllText(src.LocalPath)
            let action (source:string) (searchText:string, replaceText:string) = 
               source.Replace(searchText, replaceText)
            let final = List.fold action original searches
            let _ = File.WriteAllText(outputAbsPath, final)
            return! workingMarkdownDoc outputAbsPath
        }


    let findReplace (searches:SearchList)
                    (src:MarkdownDoc) : DocMonad<'res,MarkdownDoc> = 
        findReplaceAs searches src.LocalPath src
