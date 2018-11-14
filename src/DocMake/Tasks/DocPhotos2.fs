// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

[<RequireQualifiedAccess>]
module DocMake.Tasks.DocPhotos2


open System.IO
open System.Text.RegularExpressions
open Microsoft.Office.Interop

open MarkdownDoc
open MarkdownDoc.Pandoc

open DocMake.Base.FakeLike
open DocMake.Builder.BuildMonad
open DocMake.Builder.Document
open DocMake.Builder.Basis



// TODO 
// For simplicity of the API we should expect Photos to be copied and shrunk.



let page1 (title:string) (imagePath:string) : Markdown = 
    let imageName = System.IO.Path.GetFileNameWithoutExtension imagePath
    concat [ h1 (text title)
           ; tile <| nbsp       // should be Markdown...
           ; tile <| inlineImage (text " ") imagePath None
           ; tile <| text imageName
           ]

let pagesRest (title:string) (imagePath:string) : Markdown = 
    let imageName = System.IO.Path.GetFileNameWithoutExtension imagePath
    concat [ openxmlPagebreak
           ; h2 (text title)
           ; tile <| nbsp       // should be Markdown...
           ; tile <| inlineImage (text " ") imagePath None
           ; tile <| text imageName
           ]

let private makePhotoDocMd (title:string) (imagePaths: string list) : Markdown = 
    match imagePaths with
    | x :: xs -> 
        let rest = List.map (pagesRest title) xs
        concat (page1 title x :: rest)
    | [] -> h1 (text title)



//// TODO inputPaths should be paired with an optional rename procedure
//let private photoDocImpl (getHandle:'res-> Word.Application) (opts:DocPhotosOptions) (inputPaths:string list) : BuildMonad<'res,WordDoc> =
//    let docProc (jpegFolder:string) : DocOutput<unit>  = 
//        let jpegs = getJPEGs1 jpegFolder
//        let stepFun = if opts.ShowFileName then stepWithLabel else stepWithoutLabel
//        docOutput { 
//            do! addTitle opts.DocTitle
//            do! insertPhotos stepFun jpegs
//            }

//    buildMonad { 
//        do! mapMz (fun jpg -> copyJPEGs jpg opts.CopyToSubDirectory) inputPaths
//        let! outDoc = freshDocument "docx"
//        let! app = asksU getHandle
//        let! jpegCopiesLoc = (fun d -> d </> opts.CopyToSubDirectory) <<| askWorkingDirectory ()
//        match outDoc.GetPath with
//        | None -> return zeroDocument
//        | Some outPath -> 
//            if System.IO.Directory.Exists(jpegCopiesLoc) && System.IO.Directory.GetFiles(jpegCopiesLoc).Length > 0 then
//                let _ = runDocOutput2 outPath app (docProc jpegCopiesLoc)
//                return outDoc
//            else
//                return zeroDocument
//        } 
    


//type DocPhotosApi<'res> = 
//    { DocPhotos : string -> string list -> BuildMonad<'res, WordDoc> }

//let makeAPI (getHandle:'res -> Word.Application) : DocPhotosApi<'res> = 
//    { DocPhotos = photoDocImpl getHandle }
