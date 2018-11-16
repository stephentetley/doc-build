// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


module DocBuild.JpegDoc


open DocBuild.Internal.Common
open DocBuild.Internal.ImageMagickUtils



// NOTE:
// ImageMagick extracts image info from the original file, not the 
// image handle. Therefore we need to carry the image file around 
// even though image manipulations are performed on a handle.
//
// We destructively operate on a single temp file, this avoids generating 
// a multitude of temp files, but means we never touch the original.




type JpegDoc = 
    val private JpegPath : string
    val private TempPath : string

    new (filePath:string) = 
        { JpegPath = filePath
        ; TempPath = getTempFileName filePath }

    member internal v.TempFile
        with get() : string = 
            if System.IO.File.Exists(v.TempPath) then
                v.TempPath
            else
                System.IO.File.Copy(v.JpegPath, v.TempPath)
                v.TempPath

    member v.SaveAs(outputPath: string) :JpegDoc = 
        let updatedFile = v.TempFile
        System.IO.File.Move(updatedFile, outputPath)
        new JpegDoc(filePath = outputPath)


    member v.AutoOrient() : JpegDoc = 
        autoOrient(v.TempFile)
        v

    member v.ResizeForWord() : JpegDoc = 
        optimizeForMsWord(v.TempFile)
        v

let jpegDoc (path:string) : JpegDoc = new JpegDoc (filePath = path)

