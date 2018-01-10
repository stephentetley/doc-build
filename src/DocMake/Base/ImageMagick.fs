namespace DocMake.Base

open System
open ImageMagick


module ImageMagick = 

    type Orientation = Portrait | Landscape

    let getOrientation (info:MagickImageInfo) : Orientation = 
        if info.Width > info.Height then Landscape else Portrait
    
    let makeRevisedFileName (annotation:string)  (filePath:string) : string = 
        let root = System.IO.Path.GetDirectoryName filePath
        let justfile = System.IO.Path.GetFileNameWithoutExtension filePath
        let ext  = System.IO.Path.GetExtension filePath
        let newfile = sprintf "%s.%s%s" justfile annotation ext
        IO.Path.Combine(root, newfile)


    let calculateNewPixelSize (info:MagickImageInfo) (maxima:int) : (int * int) = 
        match getOrientation info with
        | Landscape -> 
            let scaling = maxima / info.Width in (maxima, info.Height * scaling)
        | Portrait -> 
            let scaling = maxima / info.Height in (info.Width * scaling, maxima)


    let optimizeForMsWord (filePath:string) : unit = 
        use (img:MagickImage) = new MagickImage(filePath)
        let info = new MagickImageInfo(filePath)
        // let out = makeRevisedFileName "600px" filePath
        let (newWidth,newHeight) = calculateNewPixelSize info 600
        img.Density <- new Density(72.0, 72.0, DensityUnit.PixelsPerInch)
        img.Resize(new MagickGeometry(newWidth, newHeight))
        img.Write filePath