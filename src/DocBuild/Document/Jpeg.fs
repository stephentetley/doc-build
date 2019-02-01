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
    let autoOrientAs (outputAbsPath:string) (src:JpegFile) : DocMonad<'res,JpegFile> = 
        docMonad { 
            do! assertIsWorkingPath outputAbsPath
            let _ = ImageMagickPrim.imAutoOrient src.LocalPath outputAbsPath
            return! workingJpegFile outputAbsPath
        }

    /// Auto-orient overwriting the input file
    let autoOrient (src:JpegFile) : DocMonad<'res,JpegFile> = 
        autoOrientAs src.LocalPath src


    /// Save in working directory (or a child of).
    let resizeForWordAs (outputAbsPath:string) (src:JpegFile) : DocMonad<'res,JpegFile> = 
        docMonad { 
            do! assertIsWorkingPath outputAbsPath
            let _ = ImageMagickPrim.imOptimizeForMsWord src.LocalPath outputAbsPath
            return! workingJpegFile outputAbsPath
        }

    /// Resize for Word overwriting the input file
    let resizeForWord (src:JpegFile) : DocMonad<'res,JpegFile> = 
        resizeForWordAs src.LocalPath src




