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

    /// Save in working directory...
    let autoOrientAs (outputName:string) (src:JpegFile) : DocMonad<'res,JpegFile> = 
        docMonad { 
            let! outputPath = getOutputPath outputName
            let _ = ImageMagickPrim.imAutoOrient src.AbsolutePath outputPath
            let! jpeg = workingJpegFile outputName
            return jpeg
        }

    let autoOrient (src:JpegFile) : DocMonad<'res,JpegFile> = 
        autoOrientAs src.FileName src


    /// Save in working directory...
    let resizeForWordAs (outputName:string) (src:JpegFile) : DocMonad<'res,JpegFile> = 
        docMonad { 
            let! outputPath = getOutputPath outputName
            let _ = ImageMagickPrim.imOptimizeForMsWord src.AbsolutePath outputPath
            let! jpeg = workingJpegFile outputName
            return jpeg
        }

    /// Rezize for Word generating a new temp file
    let resizeForWord (src:JpegFile) : DocMonad<'res,JpegFile> = 
        resizeForWordAs src.FileName src




