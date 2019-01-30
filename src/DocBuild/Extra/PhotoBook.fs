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

        

    let private makePage1 (title:string) 
                          (imagePath:string) 
                          (imageName:string) : Markdown = 
        concat [ h1 (text title)
               ; tile <| nbsp       // should be Markdown...
               ; tile <| inlineImage (text " ") imagePath None
               ; tile <| text imageName
               ]

    let private makePageRest (title:string) 
                             (imagePath:string) 
                             (imageName:string) : Markdown = 
        concat [ openxmlPagebreak
               ; h2 (text title)
               ; tile <| nbsp       // should be Markdown...
               ; tile <| inlineImage (text " ") imagePath None
               ; tile <| text imageName
               ]


    let private photoBookMarkdown (title:string) 
                                  (imagePaths: JpegFile list) : Markdown = 
        match imagePaths with
        | x :: xs -> 
            let page1 = makePage1 title x.AbsolutePath x.Title
            let rest = xs |> List.map (fun x -> makePageRest title x.AbsolutePath x.Title) 
            concat (page1 :: rest)
        | [] -> h1 (text title)

    
    let internal copyJpegs (sourceSubFolder:string)
                           (tempSubFolder:string) : DocMonad<'res, JpegCollection> =
        let proc1 () =  
            docMonad { 
                let! xs = findAllSourceFilesMatching "*.jpg"
                let! ys = findAllSourceFilesMatching "*.jpeg"
                let! jpegs = mapM sourceJpegFile (xs @ ys) |>> Collection.fromList
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
            let mdDoc = photoBookMarkdown title (Collection.toList jpegs)
            let! outputPath = askWorkingFile outputFile
            let! _ = Markdown.saveMarkdown outputPath.AbsolutePath mdDoc
            let! mdOutput = workingMarkdownFile outputPath.AbsolutePath
            return mdOutput
        }




