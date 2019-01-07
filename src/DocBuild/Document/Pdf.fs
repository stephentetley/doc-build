// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Document


// This should support extraction / rotation via Pdftk...

[<AutoOpen>]
module Pdf = 

    open DocBuild.Base.Document
    open DocBuild.Raw.Pdftk
    open DocBuild.Raw.PdftkRotate
    
    

    type PdfFile = 
        val private PdfDoc : Document

        new (filePath:string) = 
            { PdfDoc = new Document(filePath = filePath) }

        //member v.SaveAs(outputPath: string) : unit = 
        //    if v.Updated then 
        //        System.IO.File.Move(v.TempPath, outputPath)
        //    else
        //        System.IO.File.Copy(v.SourcePath, outputPath)

        //member v.ToDocument() : Document = 
        //    if v.Updated then 
        //        document v.TempPath
        //    else
        //        document v.SourcePath

            
        //member v.RotateEmbed(options:PdftkOptions, rotations: Rotation list)  : PdfDoc = 
        //    match pdfRotateEmbed options rotations v.TempFile v.TempFile with
        //    | Choice2Of2 i when i = 0 -> v
        //    | Choice2Of2 i -> 
        //        failwithf "PdfDoc.RotateEmbed - error code %i" i
        //    | Choice1Of2 msg -> 
        //        failwithf "PdfDoc.RotateEmbed - '%s'" msg
                
        //member v.RotateExtract(options:PdftkOptions, rotations: Rotation list)  : PdfDoc = 
        //    match pdfRotateExtract options rotations v.TempFile v.TempFile with
        //    | Choice2Of2 i when i = 0 -> v
        //    | Choice2Of2 i -> 
        //        failwithf "PdfDoc.RotateEmbed - error code %i" i
        //    | Choice1Of2 msg -> 
        //        failwithf "PdfDoc.RotateEmbed - '%s'" msg

    let pdfFile (path:string) : PdfFile = new PdfFile (filePath = path)

