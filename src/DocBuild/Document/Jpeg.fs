// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Document


module Jpeg = 

    open DocBuild.Base
    open DocBuild.Base.DocMonad
    open DocBuild.Raw.ImageMagick



    // NOTE:
    // ImageMagick extracts image info from an image file, not the 
    // image handle after an image file is opened. 


    let autoOrientAs (src:JpegFile) (outfile:string) : DocMonad<JpegFile> = 
        imAutoOrient src.Path outfile |> ignore
        jpgFile outfile

    let autoOrient (src:JpegFile) : DocMonad<JpegFile> = 
        autoOrientAs src src.NextTempName



    let resizeForWordAs (src:JpegFile) (outfile:string) : DocMonad<JpegFile> = 
        imAutoOrient src.Path outfile |> ignore
        jpgFile outfile

    /// Rezize for Word generating a new temp file
    let resizeForWord (src:JpegFile) : DocMonad<JpegFile> = 
        resizeForWordAs src src.NextTempName




