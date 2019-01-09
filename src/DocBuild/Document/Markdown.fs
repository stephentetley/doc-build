// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Document


module Markdown = 

    // open MarkdownDoc
    // open MarkdownDoc.Pandoc

    open DocBuild.Base
    open DocBuild.Base.Monad

    [<Struct>]
    type MarkdownFile = 
        | MarkdownFile of Document

        member x.Path 
            with get () : FilePath =
                match x with | MarkdownFile(p) -> p.Path

        member x.NextTempName
            with get() : FilePath = 
                match x with | MarkdownFile(p) -> p.NextTempName

        //member v.ExportAsWord(options:PandocOptions, outFile:string) : WordDoc = 
        //    let command = 
        //        makePandocCommand v.Body options.DocxReferenceDoc outFile
        //    match executeProcess options.WorkingDirectory options.PandocExe command with
        //    | Choice2Of2 i when i = 0 -> 
        //        wordDoc outFile
        //    | Choice2Of2 i -> 
        //        printfn "%s" command; failwithf "PdfDoc.Save - error code %i" i
        //    | Choice1Of2 msg -> 
        //        printfn "%s" command; failwithf "PdfDoc.Save - '%s'" msg

        //member v.ExportAsWord(options:PandocOptions) : WordDoc = 
        //    let wordOut:string = System.IO.Path.ChangeExtension(v.Body, "docx")
        //    v.ExportAsWord(options = options, outFile = wordOut)

        //member v.ExportAsPdf(options:PandocOptions, outFile:string) : PdfDoc = 
        //    let wordOut:string = System.IO.Path.ChangeExtension(outFile, "docx")
        //    let wordDoc = v.ExportAsWord(options = options, outFile = wordOut)
        //    wordDoc.ExportAsPdf(quality = WordForScreen, outFile = outFile)


        //member v.ExportAsPdf(options:PandocOptions) : PdfDoc =
        //    let outFile:string = System.IO.Path.ChangeExtension(v.Body, "pdf")
        //    v.ExportAsPdf(options = options, outFile = outFile)

    let markdownDoc (path:string) : DocBuild<MarkdownFile> = 
        getDocument ".md" path |>> MarkdownFile



