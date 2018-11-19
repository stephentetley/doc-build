// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


namespace DocBuild


[<AutoOpen>]
module JpegDoc = 


    open DocBuild.Internal.CommonUtils
    open DocBuild.Internal.ImageMagickUtils



    // NOTE:
    // ImageMagick extracts image info from the original file, not the 
    // image handle. Therefore we need to carry the image file around 
    // even though image manipulations are performed on a handle.
    //
    // We destructively operate on a single temp file, this avoids generating 
    // a multitude of temp files, but means we never touch the original.




    type JpegDoc = 
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
        
        /// This is wrong as it does not allow multiple invocations 
        /// of SaveAs.
        /// If we intend getting rid of the TEMP file, it looks
        /// like we will need a destructor.
        member v.SaveAs(outputPath: string) : unit = 
            if v.Updated then 
                System.IO.File.Move(v.TempPath, outputPath)
            else
                System.IO.File.Copy(v.SourcePath, outputPath)



        member v.AutoOrient() : JpegDoc = 
            autoOrient(v.TempFile)
            v

        member v.ResizeForWord() : JpegDoc = 
            optimizeForMsWord(v.TempFile)
            v

    let jpegDoc (path:string) : JpegDoc = new JpegDoc (filePath = path)

