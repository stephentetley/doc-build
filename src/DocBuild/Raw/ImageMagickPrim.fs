// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause

namespace DocBuild.Raw


[<RequireQualifiedAccess>]
module ImageMagickPrim = 

    open System

    open ImageMagick

    open DocBuild.Base

   


    // This may get orientation "wrong" for files when the picture 
    // orientation is stored as an Exif tag.
    let getOrientation (info:MagickImageInfo) : PageOrientation = 
        if info.Width > info.Height then 
            OrientationLandscape 
        else OrientationPortrait


    // todo should have maxwidth, maxheight
    let calculateNewPixelSize (info:MagickImageInfo) (maxWidth:int, maxHeight:int) : (int * int) = 
        let getScaling (maxi:int) (current:int) : float = 
            let maxd = float maxi
            let currentd = float current
            maxd / currentd
        let scale (i:int) (factor:float) : int = int (float i * factor)
        match getOrientation info with
        | OrientationLandscape -> 
            let scaling = getScaling maxWidth info.Width 
            (maxWidth, scale info.Height scaling)
        | OrientationPortrait -> 
            let scaling = getScaling maxHeight info.Height 
            (scale info.Width scaling, maxHeight)



    let imAutoOrient (infile:string) (outfile:string) : unit = 
        use (img:MagickImage) = new MagickImage(infile)
        img.AutoOrient () // May have Exif rotation problems...
        img.Write outfile


    /// Note - this is for inclusion in a Word document.
    /// It seems it is a good size for a Markdown doc too.
    let imOptimizeForMsWord (infile:string) (outfile:string) : unit = 
        let info = new MagickImageInfo(infile)
        use (img:MagickImage) = new MagickImage(infile)
        let (newWidth,newHeight) = calculateNewPixelSize info (600,540)   // To check
        img.Density <- new Density(72.0, 72.0, DensityUnit.PixelsPerInch)
        img.Resize(new MagickGeometry(width=newWidth, height=newHeight))
        img.Write outfile





