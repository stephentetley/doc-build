module DocMake.Base.ImageMagick

open System

open Fake
open Fake.Core.Globbing.Operators

open ImageMagick


type PhotoOrientation = PhotoPortrait | PhotoLandscape

let getOrientation (info:MagickImageInfo) : PhotoOrientation = 
    if info.Width > info.Height then PhotoLandscape else PhotoPortrait
    
let makeRevisedFileName (annotation:string)  (filePath:string) : string = 
    let root = System.IO.Path.GetDirectoryName filePath
    let justfile = System.IO.Path.GetFileNameWithoutExtension filePath
    let ext  = System.IO.Path.GetExtension filePath
    let newfile = sprintf "%s.%s%s" justfile annotation ext
    IO.Path.Combine(root, newfile)


let calculateNewPixelSize (info:MagickImageInfo) (maxima:int) : (int * int) = 
    match getOrientation info with
    | PhotoLandscape -> 
        let scaling = maxima / info.Width in (maxima, info.Height * scaling)
    | PhotoPortrait -> 
        let scaling = maxima / info.Height in (info.Width * scaling, maxima)


let optimizeForMsWord (filePath:string) : unit = 
    use (img:MagickImage) = new MagickImage(filePath)
    let info = new MagickImageInfo(filePath)
    // let out = makeRevisedFileName "600px" filePath
    let (newWidth,newHeight) = calculateNewPixelSize info 600
    img.Density <- new Density(72.0, 72.0, DensityUnit.PixelsPerInch)
    img.Resize(new MagickGeometry(newWidth, newHeight))
    img.Write filePath



let optimizePhotos (jpegFolderPath:string) : unit =
    let (jpegFiles :string list) = !! (jpegFolderPath @@ "*.jpg") |> Seq.toList
    List.iter optimizeForMsWord jpegFiles
    

