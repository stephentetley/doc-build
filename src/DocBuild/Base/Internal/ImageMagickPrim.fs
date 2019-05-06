// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause

namespace DocBuild.Base.Internal


[<RequireQualifiedAccess>]
module ImageMagickPrim = 

    open System

    open ImageMagick

      
    let isLandscape (info:MagickImageInfo) : bool = 
        info.Width > info.Height

    let isPortrait (info:MagickImageInfo) : bool = 
        info.Height > info.Width



    // todo should have maxwidth, maxheight
    let calculateNewPixelSize (info:MagickImageInfo) 
                              (maxWidth:int, maxHeight:int) : (int * int) = 
        let getScaling (maxi:int) (current:int) : float = 
            let maxd = float maxi
            let currentd = float current
            maxd / currentd
        let scale (i:int) (factor:float) : int = int (float i * factor)
        if isLandscape info then
            // Lanscape
            let scaling = getScaling maxWidth info.Width 
            (maxWidth, scale info.Height scaling)
        else
            // Portrait
            let scaling = getScaling maxHeight info.Height 
            (scale info.Width scaling, maxHeight)



    let imAutoOrient (infile:string) (outfile:string) : Result<unit, string> = 
        try 
            use (img:MagickImage) = new MagickImage(infile)
            img.AutoOrient () // May have Exif rotation problems...
            img.Write outfile
            Ok ()
        with
        | exn -> Error (sprintf "imAutoOrient: %s" exn.Message)


    /// Note - this is for inclusion in a Word document.
    /// It seems it is a good size for a Markdown doc too.
    let imOptimizeForMsWord (infile:string) (outfile:string) : Result<unit, string> = 
        try 
            let info = new MagickImageInfo(infile)
            use (img:MagickImage) = new MagickImage(infile)
            let (newWidth,newHeight) = calculateNewPixelSize info (600,540)   // To check
            img.Density <- new Density(72.0, 72.0, DensityUnit.PixelsPerInch)
            img.Resize(new MagickGeometry(width=newWidth, height=newHeight))
            img.Write outfile
            Ok ()
        with
        | exn -> Error (sprintf "imAutoOrient: %s" exn.Message)




