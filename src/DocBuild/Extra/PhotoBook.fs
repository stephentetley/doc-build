// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause

namespace DocBuild.Extra

module PhotoBook = 

    open MarkdownDoc
    open MarkdownDoc.Pandoc

    open DocBuild.Base

    open DocBuild.Document
    open DocBuild.Document.Jpeg




    let optimizeJpeg (image:JpegDoc) : DocMonad<JpegDoc, 'userRes> =
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
            let page1 = makePage1 title x.AbsolutePath x.Title
            let rest = xs |> List.map (fun x -> makePageRest title x.AbsolutePath x.Title) 
            concatMarkdown (page1 :: rest)
        | [] -> h1 (text title)

    
    let internal copyJpegs (sourceSubdirectory:string) : DocMonad<JpegCollection, 'userRes> =
        docMonad { 
            let! xs = 
                localSourceSubdirectory sourceSubdirectory <| findSourceFilesMatching "*.jpg" false
            let! ys = 
                localSourceSubdirectory sourceSubdirectory <| findSourceFilesMatching "*.jpeg" false
            let! srcJpegs = mapM getJpegDoc (xs @ ys) |>> List.sortBy (fun doc -> doc.FileName)
            let! jpegs = mapM copyToWorking srcJpegs
            return (Collection.fromList jpegs)
        }

    let internal optimizeJpegs (jpegs:JpegCollection) : DocMonad<JpegCollection, 'userRes> =
        mapM optimizeJpeg jpegs.Elements |>> Collection.fromList



    type PhotoBookConfig = 
        { Title: string
          SourceSubdirectory: string
          WorkingSubdirectory: string
          RelativeOutputName: string }

    /// Check for empty & handle missing source folder.
    let makePhotoBook (config:PhotoBookConfig) : DocMonad<MarkdownDoc option, 'userRes> =
        docMonad {
            let! jpegs = 
                localWorkingSubdirectory config.WorkingSubdirectory <|
                    optionMaybeM (copyJpegs config.SourceSubdirectory >>= optimizeJpegs)
            match jpegs with
            | None -> return None
            | Some col when col.Elements.IsEmpty -> return None
            | Some col -> 
                let mdDoc = photoBookMarkdown config.Title col.Elements
                return! Markdown.saveMarkdown config.RelativeOutputName mdDoc |>> Some
        }

    let genPhotoBook (render: MarkdownDoc -> DocMonad<PdfDoc, 'userRes>)
                     (config:PhotoBookConfig) : DocMonad<PdfDoc, 'userRes> =
        docMonad { 
            match! makePhotoBook config with
            | Some md -> return! render md
            | None -> return! docError "makePhotoBook failed"
        }

