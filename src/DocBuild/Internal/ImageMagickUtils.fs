// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

module DocBuild.Internal.ImageMagickUtils

open System

open ImageMagick


type PhotoOrientation = PhotoPortrait | PhotoLandscape


// This may get orientation "wrong" for files when the picture 
// orientation is stored as an Exif tag.
let getOrientation (info:MagickImageInfo) : PhotoOrientation = 
    if info.Width > info.Height then PhotoLandscape else PhotoPortrait
    
let makeRevisedFileName (annotation:string)  (filePath:string) : string = 
    let root = System.IO.Path.GetDirectoryName filePath
    let justfile = System.IO.Path.GetFileNameWithoutExtension filePath
    let ext  = System.IO.Path.GetExtension filePath
    let newfile = sprintf "%s.%s%s" justfile annotation ext
    IO.Path.Combine(root, newfile)

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


type IMQuery<'a> = MagickImage -> 'a
type IMTransform = MagickImage -> unit

let autoOrient : IMTransform = fun img -> img.AutoOrient ()

//let autoOrient (filePath:string) : unit = 
//    use (img:MagickImage) = new MagickImage(filePath)
//    img.AutoOrient () // May have Exif rotation problems...
//    img.Write filePath



let optimizeForMsWord (filePath:string) : unit = 
    let info = new MagickImageInfo(filePath)
    use (img:MagickImage) = new MagickImage(filePath)
    let (newWidth,newHeight) = calculateNewPixelSize info (600,540)   // To check
    img.Density <- new Density(72.0, 72.0, DensityUnit.PixelsPerInch)
    img.Resize(new MagickGeometry(width=newWidth, height=newHeight))
    img.Write filePath




//let optimizePhotos (jpegFolderPath:string) : unit =
//    let (jpegFiles :string list) = getFilesMatching jpegFolderPath "*.jpg"
//    List.iter (fun file -> autoOrient file; optimizeForMsWord file) jpegFiles
    

