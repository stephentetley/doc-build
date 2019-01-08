// Copyright (c) Stephen Tetley 2018,2019
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




    type JpegFile = 
        val private JpegDoc : Document

        new (filePath:string) = 
            { JpegDoc = new Document(filePath = filePath) }


        

        member x.SaveAs(outputPath: string) : unit =  
            x.JpegDoc.SaveAs(outputPath)



        member x.AutoOrient() : unit = 
            autoOrient x.JpegDoc.TempFile


        member x.ResizeForWord() : unit = 
            optimizeForMsWord x.JpegDoc.TempFile


    let jpegFile (path:string) : JpegFile = new JpegFile (filePath = path)

    let autoOrient (src:JpegFile) : JpegFile = 
        ignore <| src.AutoOrient() ; src

    let resizeForWord (src:JpegFile) : JpegFile = 
        ignore <| src.ResizeForWord() ; src

    let saveJpegFile (outputName:string) (doc:JpegFile) : unit = 
        doc.SaveAs(outputPath = outputName)


