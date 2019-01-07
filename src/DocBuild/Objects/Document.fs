// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

namespace DocBuild


[<AutoOpen>]
module Document = 

    open DocBuild.Base
    open DocBuild.Raw.Ghostscript

    /// Concat PDFs with Ghostscript
    /// We favour Ghostscript because it lets us lower the print 
    /// quality (and reduce the file size).

    type GhostscriptOptions = DocBuild.Raw.Ghostscript.GhostscriptOptions
    type GsPdfQuality = DocBuild.Raw.Ghostscript.GsPdfQuality

    type PdfPath = string




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
        
        static member (^^) (doc1:Document, doc2:Document) : Document = 
            new Document(paths = doc1.Body @ doc2.Body)


        member v.SaveAs(options: GhostscriptOptions, outputPath: string) : unit = 
            let command = makeGsConcatCommand options.PrintQuality outputPath v.Body
            match runGhostscript options command with
            | Choice2Of2 i when i = 0 -> ()
            | Choice2Of2 i -> 
                printfn "%s" command; failwithf "PdfDoc.Save - error code %i" i
            | Choice1Of2 msg -> 
                printfn "%s" command; failwithf "PdfDoc.Save - '%s'" msg



    let emptyDocument: Document = new Document ()

    let document (path:PdfPath) : Document = new Document (filePath = path)


    let concat (docs:Document list) = 
        let xs = List.concat (List.map (fun (d:Document) -> d.Body) docs)
        new Document(paths = xs)