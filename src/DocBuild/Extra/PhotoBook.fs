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
            let page1 = makePage1 title x.LocalPath x.Title
            let rest = xs |> List.map (fun x -> makePageRest title x.LocalPath x.Title) 
            concat (page1 :: rest)
        | [] -> h1 (text title)

    
    let internal copyJpegs (sourceSubFolder:string)
                           (tempSubFolder:string) : DocMonad<'res, JpegCollection> =
        let proc1 () =  
            let action (fullPath:string) = sourceJpegFile (FileInfo(fullPath).Name)
            docMonad { 
                let! xs = findAllSourceFilesMatching "*.jpg" false
                let! ys = findAllSourceFilesMatching "*.jpeg" false
                printfn "copyJpegs - here 1"
                List.iter (printfn "%s") xs 
                let! jpegs = mapM action (xs @ ys) |>> Collection.fromList
                printfn "copyJpegs - here 2"
                return jpegs
            }
        printfn "Source: '%s'" sourceSubFolder
        printfn "Working: '%s'" sourceSubFolder
        localSubDirectory tempSubFolder 
            << childSourceDirectory sourceSubFolder <| proc1 ()

    let makePhotoBook (title:string) 
                      (sourceSubFolder:string) 
                      (tempSubFolder:string)
                      (outputName:string) : DocMonad<'res,MarkdownFile> =
        docMonad {
            let! jpegs = copyJpegs sourceSubFolder tempSubFolder >>= Collection.mapM optimizeJpeg
            let mdDoc = photoBookMarkdown title (Collection.toList jpegs)
            let! outputPath = askWorkingFile outputName
            let! _ = Markdown.saveMarkdown outputPath.LocalPath mdDoc
            let! mdOutput = workingMarkdownFile outputPath.LocalPath
            return mdOutput
        }




