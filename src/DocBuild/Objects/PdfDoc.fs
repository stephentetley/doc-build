// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


namespace DocBuild


// This should support extraction / rotation via Pdftk...

[<AutoOpen>]
module PdfDoc = 

    open DocBuild.Base
    open DocBuild.Raw.PdftkRotate
    
    

    type PdfDoc = 
        val private SourcePath : string
        val private TempPath : string

        new (filePath:string) = 
            { SourcePath = filePath
            ; TempPath = Temp.getTempFileName filePath }

        member internal v.TempFile
            with get() : string = 
                if System.IO.File.Exists(v.TempPath) then
                    v.TempPath
                else
                    System.IO.File.Copy(v.SourcePath, v.TempPath)
                    v.TempPath
    
        member internal v.Updated 
            with get() : bool = System.IO.File.Exists(v.TempPath)


        member v.SaveAs(outputPath: string) : unit = 
            if v.Updated then 
                System.IO.File.Move(v.TempPath, outputPath)
            else
                System.IO.File.Copy(v.SourcePath, outputPath)

        member v.ToDocument() : Document = 
            if v.Updated then 
                document v.TempPath
            else
                document v.SourcePath

            
        member v.RotateEmbed(options:PdftkOptions, rotations: Rotation list)  : PdfDoc = 
            match pdfRotateEmbed options rotations v.TempFile v.TempFile with
            | Choice2Of2 i when i = 0 -> v
            | Choice2Of2 i -> 
                failwithf "PdfDoc.RotateEmbed - error code %i" i
            | Choice1Of2 msg -> 
                failwithf "PdfDoc.RotateEmbed - '%s'" msg
                
        member v.RotateExtract(options:PdftkOptions, rotations: Rotation list)  : PdfDoc = 
            match pdfRotateExtract options rotations v.TempFile v.TempFile with
            | Choice2Of2 i when i = 0 -> v
            | Choice2Of2 i -> 
                failwithf "PdfDoc.RotateEmbed - error code %i" i
            | Choice1Of2 msg -> 
                failwithf "PdfDoc.RotateEmbed - '%s'" msg

    let pdfDoc (path:string) : PdfDoc = new PdfDoc (filePath = path)

    let toDocument (pdfDoc:PdfDoc) : Document = 
        pdfDoc.ToDocument()

    let ( *^^ ) (doc:Document) (pdf:PdfDoc) : Document = 
        doc ^^ pdf.ToDocument()