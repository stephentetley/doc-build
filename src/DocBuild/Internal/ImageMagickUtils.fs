// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

namespace DocBuild.Internal.ImageMagickUtils

[<AutoOpen>]
module ImageMagickUtils = 

    open System

    open ImageMagick


    type PhotoOrientation = PhotoPortrait | PhotoLandscape


    // This may get orientation "wrong" for files when the picture 
    // orientation is stored as an Exif tag.
    let getOrientation (info:MagickImageInfo) : PhotoOrientation = 
        if info.Width > info.Height then PhotoLandscape else PhotoPortrait


    // todo should have maxwidth, maxheight
    let calculateNewPixelSize (info:MagickImageInfo) (maxWidth:int, maxHeight:int) : (int * int) = 
        let getScaling (maxi:int) (current:int) : float = 
            let maxd = float maxi
            let currentd = float current
            maxd / currentd
        let scale (i:int) (factor:float) : int = int (float i * factor)
        match getOrientation info with
        | PhotoLandscape -> 
            let scaling = getScaling maxWidth info.Width 
            (maxWidth, scale info.Height scaling)
        | PhotoPortrait -> 
            let scaling = getScaling maxHeight info.Height 
            (scale info.Width scaling, maxHeight)



    let autoOrient (filePath:string) : unit = 
        use (img:MagickImage) = new MagickImage(filePath)
        img.AutoOrient () // May have Exif rotation problems...
        img.Write filePath



    let optimizeForMsWord (filePath:string) : unit = 
        let info = new MagickImageInfo(filePath)
        use (img:MagickImage) = new MagickImage(filePath)
        let (newWidth,newHeight) = calculateNewPixelSize info (600,540)   // To check
        img.Density <- new Density(72.0, 72.0, DensityUnit.PixelsPerInch)
        img.Resize(new MagickGeometry(width=newWidth, height=newHeight))
        img.Write filePath





