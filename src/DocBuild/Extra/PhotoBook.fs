// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause

namespace DocBuild.Extra

module PhotoBook = 

    open System.IO


    open MarkdownDoc
    open MarkdownDoc.Pandoc

    open DocBuild.Base
    open DocBuild.Base.DocMonad
    open DocBuild.Base.DocMonadOperators
    // open DocBuild.Base.FakeLike


    open DocBuild.Document.Jpeg
    open DocBuild.Document




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

    
    let internal copyJpegs (sourceSubFolder:string)
                           (tempSubFolder:string) : DocMonad<'res, JpegCollection> =
        let proc1 () =  
            docMonad { 
                let! xs = findAllSourceFilesMatching "*.jpg"
                let! ys = findAllSourceFilesMatching "*.jpeg"
                let! col1 = mapM getJpegFile (xs @ ys) |>> Collection.fromList
                let! jpegs = copyCollectionToWorking col1
                return jpegs
            }
        localSubDirectory tempSubFolder 
            << childSourceDirectory sourceSubFolder <| proc1 ()

    let makePhotoBook (title:string) 
                      (sourceSubFolder:string) 
                      (tempSubFolder:string)
                      (outputFile:string) : DocMonad<'res,MarkdownFile> =
        docMonad {
            let! jpegs = copyJpegs sourceSubFolder tempSubFolder >>= Collection.mapM optimizeJpeg
            let jpegPaths = Collection.toList jpegs |> List.map (fun jpg1 -> jpg1.Path)
            let mdDoc = photoBookMarkdown title jpegPaths
            let! outputPath = askWorkingFile outputFile
            let! _ = Markdown.saveMarkdown outputPath mdDoc
            let! mdOutput = getMarkdownFile outputPath
            return mdOutput
        }




