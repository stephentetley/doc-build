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

    /// Save in working directory (or a child of).
    let autoOrientAs (outputAbsPath:string) (src:JpegDoc) : DocMonad<'res,JpegDoc> = 
        docMonad { 
            do! assertIsWorkingPath outputAbsPath
            let _ = ImageMagickPrim.imAutoOrient src.LocalPath outputAbsPath
            return! workingJpegDoc outputAbsPath
        }

    /// Auto-orient overwriting the input file
    let autoOrient (src:JpegDoc) : DocMonad<'res,JpegDoc> = 
        autoOrientAs src.LocalPath src


    /// Save in working directory (or a child of).
    let resizeForWordAs (outputAbsPath:string) (src:JpegDoc) : DocMonad<'res,JpegDoc> = 
        docMonad { 
            do! assertIsWorkingPath outputAbsPath
            let _ = ImageMagickPrim.imOptimizeForMsWord src.LocalPath outputAbsPath
            return! workingJpegDoc outputAbsPath
        }

    /// Resize for Word overwriting the input file
    let resizeForWord (src:JpegDoc) : DocMonad<'res,JpegDoc> = 
        resizeForWordAs src.LocalPath src




