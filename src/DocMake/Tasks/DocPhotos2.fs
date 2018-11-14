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
open DocMake.Builder.PandocRunner



// Note 
// For simplicity of the API we should expect Photos to be copied and shrunk.



let private makePage1 (title:string) (imagePath:string) : Markdown = 
    let imageName = System.IO.Path.GetFileNameWithoutExtension imagePath
    concat [ h1 (text title)
           ; tile <| nbsp       // should be Markdown...
           ; tile <| inlineImage (text " ") imagePath None
           ; tile <| text imageName
           ]

let private makePageRest (title:string) (imagePath:string) : Markdown = 
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
        let page1 = makePage1 title x
        let rest = List.map (makePageRest title) xs
        concat (page1 :: rest)
    | [] -> h1 (text title)



let makePhotoDocx (title:string) (imagePaths: string list) 
                    (outputDocx:string) : PandocRunner<WordDoc> = 
    let mdDoc = makePhotoDocMd title imagePaths
    generateDocx mdDoc outputDocx []



