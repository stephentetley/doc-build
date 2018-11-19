// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


namespace DocBuild


[<AutoOpen>]
module MarkdownDoc = 

    open MarkdownDoc

    open DocBuild.Internal.RunProcess
    open DocBuild


    // pandoc -f markdown -t docx+table_captions <INFILE> --reference-doc=<CUSTOM_REF> -s -o <OUTFILE>

    let private makePandocCommand (inFile:string) (customRef:string) 
                                    (outFile:string) : string = 
        sprintf "-f markdown -t docx+table_captions \"%s\" --reference-doc=\"%s\" -s -o \"%s\""
                    inFile customRef outFile

    type PandocOptions = 
        { WorkingDirectory: string 
          PandocExe: string 
          DocxReferenceDoc: string
        }

    type MarkdownDoc = 
    
        // Design note - path or Markdown (from MarkdownDoc)
        // Path is consistent with other objects and we can always render 
        // ``Markdown`` to text.

        val private MarkdownPath : string

        new (filePath:string) = 
            { MarkdownPath = filePath }

        new (markdown:Markdown, filePath:string) = 
            markdown.Save(filePath); { MarkdownPath = filePath }


        member internal v.Body 
            with get() : string = v.MarkdownPath

        member v.ExportAsWord(options:PandocOptions, outFile:string) : WordDoc = 
            let command = 
                makePandocCommand v.Body options.DocxReferenceDoc outFile
            match executeProcess options.WorkingDirectory options.PandocExe command with
            | Choice2Of2 i when i = 0 -> 
                wordDoc outFile
            | Choice2Of2 i -> 
                printfn "%s" command; failwithf "PdfDoc.Save - error code %i" i
            | Choice1Of2 msg -> 
                printfn "%s" command; failwithf "PdfDoc.Save - '%s'" msg

        member v.ExportAsWord(options:PandocOptions) : WordDoc = 
            let wordOut:string = System.IO.Path.ChangeExtension(v.Body, "docx")
            v.ExportAsWord(options = options, outFile = wordOut)

        member v.ExportAsPdf(options:PandocOptions, outFile:string) : PdfDoc = 
            let wordOut:string = System.IO.Path.ChangeExtension(outFile, "docx")
            let wordDoc = v.ExportAsWord(options = options, outFile = wordOut)
            wordDoc.ExportAsPdf(quality = WordForScreen, outFile = outFile)


        member v.ExportAsPdf(options:PandocOptions) : PdfDoc =
            let outFile:string = System.IO.Path.ChangeExtension(v.Body, "pdf")
            v.ExportAsPdf(options = options, outFile = outFile)

    let markdownDoc (path:string) : MarkdownDoc = new MarkdownDoc (filePath = path)

    let markdownDoc2 (path:string) (markdown:Markdown) : MarkdownDoc = 
        new MarkdownDoc (markdown = markdown, filePath = path)

