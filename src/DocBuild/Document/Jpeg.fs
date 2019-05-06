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
    let autoOrientAs (outputRelName:string) (src:JpegDoc) : DocMonad<'userRes,JpegDoc> = 
        docMonad { 
            let! outputAbsPath = extendWorkingPath outputRelName
            let! _ = 
                liftOperationResult "ImageMagick auto orient error" 
                    <| fun _ -> ImageMagickPrim.imAutoOrient src.AbsolutePath outputAbsPath
            return! getJpegDoc outputAbsPath
        }

    /// Auto-orient overwriting the input file
    let autoOrient (src:JpegDoc) : DocMonad<'userRes,JpegDoc> = 
        autoOrientAs src.FileName src


    /// Save in working directory (or a child of).
    let resizeForWordAs (outputRelName:string) (src:JpegDoc) : DocMonad<'userRes,JpegDoc> = 
        docMonad { 
            let! outputAbsPath = extendWorkingPath outputRelName
            let! _ = 
                liftOperationResult "ImageMagick resize error" 
                    <| fun _ -> ImageMagickPrim.imOptimizeForMsWord src.AbsolutePath outputAbsPath
            return! getJpegDoc outputAbsPath
        }

    /// Resize for Word overwriting the input file
    let resizeForWord (src:JpegDoc) : DocMonad<'userRes,JpegDoc> = 
        resizeForWordAs src.FileName src




