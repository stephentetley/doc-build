// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

[<RequireQualifiedAccess>]
module DocMake.Tasks.DocPhotos2


open System.IO


open MarkdownDoc
open MarkdownDoc.Pandoc

open DocBuild.MarkdownDoc


// Note 
// For simplicity of the API we should resize and AutoOrient the photos first.

// TODO - optimize images...

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

let private photoBookMarkdown (title:string) (imagePaths: string list) : Markdown = 
    match imagePaths with
    | x :: xs -> 
        let page1 = makePage1 title x
        let rest = List.map (makePageRest title) xs
        concat (page1 :: rest)
    | [] -> h1 (text title)



let makePhotoBook (title:string) (imagePaths: string list) 
                    (outFile:string) : MarkdownDoc = 
    let book = photoBookMarkdown title imagePaths
    new MarkdownDoc (markdown = book, filePath = outFile)



