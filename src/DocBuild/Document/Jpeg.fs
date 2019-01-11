// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Document


module Jpeg = 

    open DocBuild.Base
    open DocBuild.Base.DocMonad
    open DocBuild.Raw



    // NOTE:
    // ImageMagick extracts image info from an image file, not the 
    // image handle after an image file is opened. 


    let autoOrientAs (src:JpegFile) (outputFile:string) : DocMonad<'res,JpegFile> = 
        ImageMagickPrim.imAutoOrient src.Path outputFile |> ignore
        getJpegFile outputFile

    let autoOrient (src:JpegFile) : DocMonad<'res,JpegFile> = 
        autoOrientAs src src.NextTempName



    let resizeForWordAs (src:JpegFile) (outputFile:string) : DocMonad<'res,JpegFile> = 
        ImageMagickPrim.imAutoOrient src.Path outputFile |> ignore
        getJpegFile outputFile

    /// Rezize for Word generating a new temp file
    let resizeForWord (src:JpegFile) : DocMonad<'res,JpegFile> = 
        resizeForWordAs src src.NextTempName




