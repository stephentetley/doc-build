// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause

namespace DocBuild.Extra

module PhotoBook = 


    open MarkdownDoc
    open MarkdownDoc.Pandoc

    open DocBuild.Base
    open DocBuild.Base.DocMonad
    open DocBuild.Base.DocMonadOperators


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
            let page1 = makePage1 title x.LocalPath x.Title
            let rest = xs |> List.map (fun x -> makePageRest title x.LocalPath x.Title) 
            concat (page1 :: rest)
        | [] -> h1 (text title)

    
    let internal copyJpegs (sourceSubdirectory:string)
                           (workingSubdirectory:string) : DocMonad<'res, JpegCollection> =
        docMonad { 
            let! xs = 
                localSourceSubdirectory sourceSubdirectory <| findAllSourceFilesMatching "*.jpg" false
            let! ys = 
                localSourceSubdirectory sourceSubdirectory <| findAllSourceFilesMatching "*.jpeg" false
            let! jpegs = mapM (copyFileToWorkingSubdirectory workingSubdirectory) (xs @ ys)
            return (Collection.fromList jpegs)
        }


    let makePhotoBook (title:string) 
                      (sourceSubFolder:string) 
                      (tempSubFolder:string)
                      (outputName:string) : DocMonad<'res,MarkdownFile> =
        docMonad {
            let! jpegs = copyJpegs sourceSubFolder tempSubFolder >>= Collection.mapM optimizeJpeg
            let mdDoc = photoBookMarkdown title (Collection.toList jpegs)
            let! outputPath = extendWorkingPath outputName
            let! _ = Markdown.saveMarkdown outputPath mdDoc
            let! mdOutput = workingMarkdownFile outputPath
            return mdOutput
        }




