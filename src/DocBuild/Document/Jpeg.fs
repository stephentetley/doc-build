// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


namespace DocBuild.Document


module Jpeg = 

    open DocBuild.Base.Document
    open DocBuild.Raw.ImageMagick



    // NOTE:
    // ImageMagick extracts image info from an image file, not the 
    // image handle after an image file is opened. Therefore we need 
    // to carry the image file around.
    //
    // We destructively operate on a single temp file, this avoids generating 
    // a multitude of temp files, but still means we never touch the original.




    type JpegDoc = 
        val private JpegFile : Document

        new (filePath:string) = 
            { JpegFile = new Document(filePath = filePath) }


        
        /// This is wrong as it does not allow multiple invocations 
        /// of SaveAs.
        /// If we intend getting rid of the TEMP file, it looks
        /// like we will need a destructor.
        member x.SaveAs(outputPath: string) : unit =  
            x.JpegFile.SaveAs(outputPath)



        member x.AutoOrient() : unit = 
            autoOrient x.JpegFile.TempFile


        member x.ResizeForWord() : unit = 
            optimizeForMsWord x.JpegFile.TempFile


    let jpegDoc (path:string) : JpegDoc = new JpegDoc (filePath = path)

    let autoOrient (src:JpegDoc) : JpegDoc = 
        ignore <| src.AutoOrient() ; src

    let resizeForWord (src:JpegDoc) : JpegDoc = 
        ignore <| src.ResizeForWord() ; src

    let saveJpeg (outputName:string) (doc:JpegDoc) : unit = 
        doc.SaveAs(outputPath = outputName)


