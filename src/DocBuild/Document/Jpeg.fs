// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Document


module Jpeg = 

    open DocBuild.Base
    open DocBuild.Base.Monad
    open DocBuild.Raw.ImageMagick



    // NOTE:
    // ImageMagick extracts image info from an image file, not the 
    // image handle after an image file is opened. Therefore we need 
    // to carry the image file around.
    //
    // We destructively operate on a single temp file, this avoids generating 
    // a multitude of temp files, but still means we never touch the original.



    [<Struct>]
    type JpegFile = 
        | JpegFile of Document

        member x.Path 
            with get () : FilePath =
                match x with | JpegFile(p) -> p.Path

        /// ActiveFile is a mutable working copy of the original file.
        /// The original file is untouched.
        member x.NextTempName
            with get() : FilePath = 
                match x with | JpegFile(p) -> p.NextTempName



    let jpgFile (path:string) : DocBuild<JpegFile> = 
        altM (getDocument ".jpg" path) (getDocument ".jpeg" path) |>> JpegFile


    //let autoOrient (src:JpegFile) : JpegFile = 
    //    ignore <| src.AutoOrient() ; src

    //let resizeForWord (src:JpegFile) : JpegFile = 
    //    ignore <| src.ResizeForWord() ; src

    //let saveJpegFile (outputName:string) (doc:JpegFile) : unit = 
    //    doc.SaveAs(outputPath = outputName)




