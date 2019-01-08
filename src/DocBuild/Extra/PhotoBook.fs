// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause

namespace DocMake.Extra.PhotoBook

module PhotoBook = 

    open System.IO


    open MarkdownDoc
    open MarkdownDoc.Pandoc

    open DocBuild.Base.Common
    open DocBuild.Document.Markdown
    open DocBuild.Document.Jpeg

    // Note 
    // For simplicity of the API, we should resize and AutoOrient the photos first.

    // TODO - optimize images...

    let optimizeImage (imagePath:string) : string = 
        let outputname = suffixFileName "TEMP" imagePath
        (jpegFile imagePath) 
            |> autoOrient
            |> resizeForWord
            |> saveJpegFile outputname
        outputname
        

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
                        (outFile:string) : MarkdownFile =
        let newImages = List.map optimizeImage imagePaths                        
        let book = photoBookMarkdown title newImages
        ignore <| book.Save(outFile)
        new MarkdownFile (filePath = outFile)



