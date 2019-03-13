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

    let internal optimizeJpegs (jpegs:JpegCollection) : DocMonad<'res, JpegCollection> =
        mapM optimizeJpeg jpegs.Elements |>> Collection.fromList



    type PhotoBookConfig = 
        { Title: string
          SourceSubFolder: string
          WorkingSubFolder: string
          RelativeOutputName: string }

    /// Check for empty & handle missing source folder.
    let makePhotoBook (config:PhotoBookConfig) : DocMonad<'res,MarkdownDoc option> =
        docMonad {
            let! jpegs = 
                optionalM (copyJpegs config.SourceSubFolder config.WorkingSubFolder 
                            >>= optimizeJpegs)
            match jpegs with
            | None -> return None
            | Some col when col.Elements.IsEmpty -> return None
            | Some col -> 
                let mdDoc = photoBookMarkdown config.Title col.Elements
                let! outputAbsPath = extendWorkingPath config.RelativeOutputName
                let! _ = Markdown.saveMarkdown outputAbsPath mdDoc
                let! mdOutput = workingMarkdownDoc outputAbsPath
                return (Some mdOutput)
        }




