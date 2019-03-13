﻿// Copyright (c) Stephen Tetley 2018,2019
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




    let optimizeJpeg (image:JpegDoc) : DocMonad<'res,JpegDoc> =
        autoOrient image >>= resizeForWord

    let unixLikePath (path:string) : string = 
        path.Replace('\\', '/')

    let private makePage1 (title:string) 
                          (imagePath:string) 
                          (imageName:string) : Markdown = 
        concatMarkdown
            <| [ h1 (text title)
               ; nbsp       // should be Markdown...
               ; markdownText <| inlineImage "" imagePath None
               ; markdownText <| text imageName
               ]



    let private makePageRest (title:string) 
                             (imagePath:string) 
                             (imageName:string) : Markdown = 
        concatMarkdown 
            <| [ openxmlPagebreak
               ; h2 (text title)
               ; nbsp       // should be Markdown...
               ; markdownText   <| inlineImage "" imagePath None
               ; markdownText   <| text imageName
               ]


    let private photoBookMarkdown (title:string) 
                                  (imagePaths: JpegDoc list) : Markdown = 
        match imagePaths with
        | x :: xs -> 
            let page1 = makePage1 title x.LocalPath x.Title
            let rest = xs |> List.map (fun x -> makePageRest title x.LocalPath x.Title) 
            concatMarkdown (page1 :: rest)
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

    type PhotoBookConfig = 
        { Title: string
          SourceSubFolder: string
          WorkingSubFolder: string
          RelativeOutputName: string
        }

    /// TODO should we check for nonempty?
    let makePhotoBook (config:PhotoBookConfig) : DocMonad<'res,MarkdownDoc option> =
        docMonad {
            let! jpegs = 
                copyJpegs config.SourceSubFolder config.WorkingSubFolder 
                    >>= Collection.mapM optimizeJpeg
            if jpegs.Elements.IsEmpty then 
                return None
            else
                let mdDoc = photoBookMarkdown config.Title (Collection.toList jpegs)
                let! outputAbsPath = extendWorkingPath config.RelativeOutputName
                let! _ = Markdown.saveMarkdown outputAbsPath mdDoc
                let! mdOutput = workingMarkdownDoc outputAbsPath
                return (Some mdOutput)
        }




