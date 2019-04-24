// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause

namespace DocBuild.Extra

module PhotoBook = 

    open MarkdownDoc
    open MarkdownDoc.Pandoc

    open DocBuild.Base
    open DocBuild.Base.DocMonad

    open DocBuild.Document
    open DocBuild.Document.Jpeg
    open DocBuild.Document.Markdown




    let optimizeJpeg (image:JpegDoc) : DocMonad<'userRes,JpegDoc> =
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

    
    let internal copyJpegs (sourceSubdirectory:string)
                           (workingSubdirectory:string) : DocMonad<'userRes, JpegCollection> =
        docMonad { 
            let! xs = 
                localSourceSubdirectory sourceSubdirectory <| findAllSourceFilesMatching "*.jpg" false
            let! ys = 
                localSourceSubdirectory sourceSubdirectory <| findAllSourceFilesMatching "*.jpeg" false
            let! srcJpegs = mapM getJpegDoc (xs @ ys)
            let! jpegs = 
                localWorkingSubdirectory workingSubdirectory (mapM copyToWorking srcJpegs)
            return (Collection.fromList jpegs)
        }

    let internal optimizeJpegs (jpegs:JpegCollection) : DocMonad<'userRes, JpegCollection> =
        mapM optimizeJpeg jpegs.Elements |>> Collection.fromList



    type PhotoBookConfig = 
        { Title: string
          SourceSubFolder: string
          WorkingSubFolder: string
          RelativeOutputName: string }

    /// Check for empty & handle missing source folder.
    let makePhotoBook (config:PhotoBookConfig) : DocMonad<'userRes,MarkdownDoc option> =
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
                return! getWorkingMarkdownDoc outputAbsPath |>> Some
        }

    let genPhotoBook (render: MarkdownDoc -> DocMonad<'userRes,PdfDoc>)
                     (config:PhotoBookConfig) : DocMonad<'userRes, PdfDoc option> =
        makePhotoBook config >>= fun opt -> 
        match opt with
        | Some md -> render md |>> Some
        | None -> mreturn None

