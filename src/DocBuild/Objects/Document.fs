// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


namespace DocBuild



[<AutoOpen>]
module Document = 

    open DocBuild.Base
    open DocBuild.Internal.RunProcess

    /// Concat PDFs with Ghostscript
    /// We favour Ghostscript because it lets us lower the print 
    /// quality (and reduce the file size).


 
    let private gsOptions (quality:GsPdfSettings) : string =
        match quality.PrintSetting with
        | "" -> @"-dBATCH -dNOPAUSE -q -sDEVICE=pdfwrite"
        | ss -> sprintf @"-dBATCH -dNOPAUSE -q -sDEVICE=pdfwrite -dPDFSETTINGS=%s" ss

    let private gsOutputFile (fileName:string) : string = 
        sprintf "-sOutputFile=\"%s\"" fileName

    let private gsInputFile (fileName:string) : string = sprintf "\"%s\"" fileName


    /// Apparently we cannot send multiline commands to execProcess.
    let private makeGsCommand (quality:GsPdfSettings) (outputFile:string) (inputFiles: string list) : string = 
        let line1 = gsOptions quality + " " + gsOutputFile outputFile
        let rest = List.map gsInputFile inputFiles
        String.concat " " (line1 :: rest)



    type PdfPath = string



    let private runGhostscript (options:GhostscriptOptions) (command:string) : Choice<string,int> = 
        executeProcess options.WorkingDirectory options.GhostscriptExe command

    /// A PdfDoc is actually a list of Pdf files that are rendered 
    /// to a single document with Ghostscript.
    /// This means we have monodial concatenation.
    type Document = 
        val private Documents : PdfPath list

        new () = 
            { Documents = [] }

        new (filePath:PdfPath) = 
            { Documents = [filePath] }

        internal new (paths:PdfPath list ) = 
            { Documents = paths }


        member internal v.Body 
            with get() : PdfPath list = v.Documents

        member v.SaveAs(options: GhostscriptOptions, outputPath: string) : unit = 
            let command = makeGsCommand options.PrintQuality outputPath v.Body
            match runGhostscript options command with
            | Choice2Of2 i when i = 0 -> ()
            | Choice2Of2 i -> 
                printfn "%s" command; failwithf "PdfDoc.Save - error code %i" i
            | Choice1Of2 msg -> 
                printfn "%s" command; failwithf "PdfDoc.Save - '%s'" msg



    let emptyDocument: Document = new Document ()

    let document (path:PdfPath) : Document = new Document (filePath = path)

    let (^^) (x:Document) (y:Document) : Document = 
        new Document(paths = x.Body @ y.Body)

    let concat (docs:Document list) = 
        let xs = List.concat (List.map (fun (d:Document) -> d.Body) docs)
        new Document(paths = xs)