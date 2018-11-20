// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


namespace DocBuild


// This should support extraction / rotation via Pdftk...

[<AutoOpen>]
module PdfDoc = 

    open DocBuild.Internal.CommonUtils

    type PdfDoc = 
        val private SourcePath : string
        val private TempPath : string

        new (filePath:string) = 
            { SourcePath = filePath
            ; TempPath = getTempFileName filePath }

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


    let pdfDoc (path:string) : PdfDoc = new PdfDoc (filePath = path)

