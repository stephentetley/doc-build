// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

namespace DocMake.Extras.PhotoBook

module PhotoBook = 

    open System.IO


    open MarkdownDoc
    open MarkdownDoc.Pandoc

    open DocBuild.Base
    open DocBuild.MarkdownDoc
    open DocBuild.JpegDoc

    // Note 
    // For simplicity of the API we should resize and AutoOrient the photos first.

    // TODO - optimize images...

    let optimizeImage (imagePath:string) : string = 
        let outputname = suffixFileName "TEMP" imagePath
        (jpegDoc imagePath).AutoOrient().ResizeForWord().SaveAs(outputname)
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
                        (outFile:string) : MarkdownDoc =
        let newImages = List.map optimizeImage imagePaths                        
        let book = photoBookMarkdown title newImages
        new MarkdownDoc (markdown = book, filePath = outFile)



