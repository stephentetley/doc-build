// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Document


// This should support extraction / rotation via Pdftk...

[<AutoOpen>]
module Pdf = 

    open DocBuild.Base.Shell
    open DocBuild.Base.Document
    open DocBuild.Raw.Pdftk
    open DocBuild.Raw.PdftkRotate
    
    

    type PdfFile = 
        val private PdfDoc : Document

        new (filePath:string) = 
            { PdfDoc = new Document(filePath = filePath) }

        member x.SaveAs(outputPath: string) : unit =  
            x.PdfDoc.SaveAs(outputPath)



            
        member x.RotateEmbed(options:PdftkOptions, rotations: Rotation list)  : unit = 
            match pdfRotateEmbed options rotations x.PdfDoc.TempFile x.PdfDoc.TempFile with
            | ProcSuccess -> ()
            | ProcErrorCode i -> 
                failwithf "PdfDoc.RotateEmbed - error code %i" i
            | ProcErrorMessage msg -> 
                failwithf "PdfDoc.RotateEmbed - '%s'" msg
                
        member x.RotateExtract(options:PdftkOptions, rotations: Rotation list)  : unit = 
            match pdfRotateExtract options rotations x.PdfDoc.TempFile x.PdfDoc.TempFile with
            | ProcSuccess -> ()
            | ProcErrorCode i -> 
                failwithf "PdfDoc.RotateEmbed - error code %i" i
            | ProcErrorMessage msg -> 
                failwithf "PdfDoc.RotateEmbed - '%s'" msg

    let pdfFile (path:string) : PdfFile = new PdfFile (filePath = path)

