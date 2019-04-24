// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Document


module Jpeg = 

    open DocBuild.Base
    open DocBuild.Base.Internal


    // NOTE:
    // ImageMagick extracts image info from an image file, not the 
    // image handle after an image file is opened. 

    /// Save in working directory (or a child of).
    let autoOrientAs (outputAbsPath:string) (src:JpegDoc) : DocMonad<'userRes,JpegDoc> = 
        docMonad { 
            do! assertIsWorkingPath outputAbsPath
            let _ = ImageMagickPrim.imAutoOrient src.AbsolutePath outputAbsPath
            return! getWorkingJpegDoc outputAbsPath
        }

    /// Auto-orient overwriting the input file
    let autoOrient (src:JpegDoc) : DocMonad<'userRes,JpegDoc> = 
        autoOrientAs src.AbsolutePath src


    /// Save in working directory (or a child of).
    let resizeForWordAs (outputAbsPath:string) (src:JpegDoc) : DocMonad<'userRes,JpegDoc> = 
        docMonad { 
            do! assertIsWorkingPath outputAbsPath
            let _ = ImageMagickPrim.imOptimizeForMsWord src.AbsolutePath outputAbsPath
            return! getWorkingJpegDoc outputAbsPath
        }

    /// Resize for Word overwriting the input file
    let resizeForWord (src:JpegDoc) : DocMonad<'userRes,JpegDoc> = 
        resizeForWordAs src.AbsolutePath src




