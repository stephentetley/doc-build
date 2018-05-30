module DocMake.Base.ImageMagickUtils

open System

open Fake
open Fake.Core.Globbing.Operators

open ImageMagick


type PhotoOrientation = PhotoPortrait | PhotoLandscape


// This seems to be getting orientation wrong, this is either
// a problem with Magick.NET or the jpegs under test are mis-orientated.
// Need to update Magick.NET...
let getOrientation (info:MagickImageInfo) : PhotoOrientation = 
    printfn "Orientation: w=%i, h=%i, %s" info.Width info.Height info.FileName 
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
        printfn "Landscape: scaleBy=%f" scaling
        (maxWidth, scale info.Height scaling)
    | PhotoPortrait -> 
        let scaling = getScaling maxHeight info.Height 
        printfn "Portrait: scaleBy=%f" scaling
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
    printfn "img => (w:%i, h:%i)" newWidth newHeight
    img.Resize(new MagickGeometry(width=newWidth, height=newHeight))
    img.Write filePath




let optimizePhotos (jpegFolderPath:string) : unit =
    let (jpegFiles :string list) = !! (jpegFolderPath @@ "*.jpg") |> Seq.toList
    List.iter (fun file -> autoOrient file; optimizeForMsWord file) jpegFiles
    

