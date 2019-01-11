// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause

namespace DocMake.Extra.PhotoBook

module PhotoBook = 

    open System.IO


    open MarkdownDoc
    open MarkdownDoc.Pandoc

    open DocBuild.Base.DocMonad
    open DocBuild.Base.DocMonadOperators
    open DocBuild.Base.FakeLike


    open DocBuild.Document.Jpeg
    open DocBuild.Base.Document




    let optimizeJpeg (image:JpegFile) : DocMonad<'res,JpegFile> =
        autoOrient image >>= resizeForWord

        

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

    let private photoBookMarkdown (title:string) 
                                  (imagePaths: string list) : Markdown = 
        match imagePaths with
        | x :: xs -> 
            let page1 = makePage1 title x
            let rest = List.map (makePageRest title) xs
            concat (page1 :: rest)
        | [] -> h1 (text title)

    
    let internal getOptimizedJpegs 
                    (imageFolder:string) : DocMonad<'res,JpegFile list> =
        docMonad { 
            let xs = findAllMatchingFiles "*.jpg" imageFolder
            let ys = findAllMatchingFiles "*.jpeg" imageFolder
            let! jpegs = mapM (getJpegFile >=> optimizeJpeg) (xs @ ys)
            return jpegs
        }

    let makePhotoBook (title:string) 
                      (imageFolder: string) 
                      (outputFile:string) : DocMonad<'res,MarkdownFile> =
        docMonad {
            let! jpegs = getOptimizedJpegs imageFolder
            let jpegPaths = jpegs |> List.map (fun jpg1 -> jpg1.Path)
            let mdDoc = photoBookMarkdown title jpegPaths
            do mdDoc.Save(outputFile)
            let! mdOutput = getMarkdownFile outputFile
            return mdOutput
        }




