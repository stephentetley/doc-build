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
    let autoOrientAs (outputRelName : string) 
                     (source : JpegDoc) : DocMonad<JpegDoc, 'userRes> = 
        docMonad { 
            let! outputAbsPath = extendWorkingPath outputRelName
            let! sourcePath = getDocumentPath source
            let! _ = 
                liftOperationResult "ImageMagick auto orient error" 
                    <| fun _ -> ImageMagickPrim.imAutoOrient sourcePath outputAbsPath
            return! getJpegDoc outputAbsPath
        }

    /// Auto-orient overwriting the input file
    let autoOrient (source:JpegDoc) : DocMonad<JpegDoc, 'userRes> = 
        docMonad { 
            let! sourceName = getDocumentFileName source
            return! autoOrientAs sourceName source
        }


    /// Save in working directory (or a child of).
    let resizeForWordAs (outputRelName : string) 
                        (source : JpegDoc) : DocMonad<JpegDoc, 'userRes> = 
        docMonad { 
            let! outputAbsPath = extendWorkingPath outputRelName
            let! sourcePath = getDocumentPath source
            let! _ = 
                liftOperationResult "ImageMagick resize error" 
                    <| fun _ -> ImageMagickPrim.imOptimizeForMsWord sourcePath outputAbsPath
            return! getJpegDoc outputAbsPath
        }

    /// Resize for Word overwriting the input file
    let resizeForWord (source : JpegDoc) : DocMonad<JpegDoc, 'userRes> = 
        docMonad { 
            let! sourceName = getDocumentFileName source
            return! resizeForWordAs sourceName source
        }




