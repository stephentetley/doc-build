// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


#r "netstandard"

// Office deps
#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.Word"
#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Excel\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.Excel"
#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.PowerPoint\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.PowerPoint"
#I @"C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c"
#r "office"

// ImageMagick
#I @"C:\Users\stephen\.nuget\packages\Magick.NET-Q8-AnyCPU\7.9.2\lib\netstandard20"
#r @"Magick.NET-Q8-AnyCPU.dll"

// MarkdownDoc (not on nuget.org)
#I @"C:\Users\stephen\.nuget\packages\markdowndoc\1.0.0\lib\netstandard2.0"
#r @"MarkdownDoc.dll"


#load "..\src\DocBuild\Base\Common.fs"
#load "..\src\DocBuild\Base\Shell.fs"
#load "..\src\DocBuild\Base\DocMonad.fs"
#load "..\src\DocBuild\Base\DocMonadOperators.fs"
#load "..\src\DocBuild\Base\Document.fs"
#load "..\src\DocBuild\Base\Collection.fs"
#load "..\src\DocBuild\Base\FakeLike.fs"
#load "..\src\DocBuild\Base\FileIO.fs"
#load "..\src\DocBuild\Raw\GhostscriptPrim.fs"
#load "..\src\DocBuild\Raw\PandocPrim.fs"
#load "..\src\DocBuild\Raw\PdftkPrim.fs"
#load "..\src\DocBuild\Raw\ImageMagickPrim.fs"
#load "..\src\DocBuild\Document\Pdf.fs"
#load "..\src\DocBuild\Document\Jpeg.fs"
#load "..\src\DocBuild\Document\Markdown.fs"
#load "..\src\DocBuild\Extra\PhotoBook.fs"

#load "..\src-msoffice\DocBuild\Office\Internal\Utils.fs"
#load "..\src-msoffice\DocBuild\Office\Internal\WordPrim.fs"
#load "..\src-msoffice\DocBuild\Office\Internal\ExcelPrim.fs"
#load "..\src-msoffice\DocBuild\Office\Internal\PowerPointPrim.fs"
#load "..\src-msoffice\DocBuild\Office\WordFile.fs"
#load "..\src-msoffice\DocBuild\Office\ExcelFile.fs"
#load "..\src-msoffice\DocBuild\Office\PowerPointFile.fs"

open DocBuild.Base
open DocBuild.Base.DocMonad
open DocBuild.Base.DocMonadOperators
open DocBuild.Document
open DocBuild.Extra.PhotoBook
open DocBuild.Office

#load "Coversheet.fs"
open Coversheet


let inputRoot   = @"G:\work\Projects\events2\final-docs\input\CSO_SPS"
let outputRoot  = @"G:\work\Projects\events2\final-docs\output\CSO_SPS"

let (docxCustomReference:string) = @"custom-reference1.docx"

type DocMonadWord<'a> = DocMonad<WordFile.WordHandle, 'a>

let WindowsEnv : BuilderEnv = 
    { WorkingDirectory = @"G:\work\Projects\events2\final-docs\output\CSO_SPS"
      SourceDirectory =  @"G:\work\Projects\events2\final-docs\input\CSO_SPS"
      IncludeDirectory = @"G:\work\Projects\events2\final-docs\input\include"
      GhostscriptExe = @"C:\programs\gs\gs9.15\bin\gswin64c.exe"
      PdftkExe = @"pdftk"
      PandocExe = @"pandoc" }


let getSiteName (folderName:string) : string = 
    folderName.Replace('_', '/')


let getSaiNumber (siteName:string) : DocMonad<'res,string> = 
    dreturn "SAI00001234"       // TEMP

let commonSubFolder (subFolderName:string) 
                    (ma:DocMonad<'res,'a>) : DocMonad<'res,'a> = 
    localSubDirectory subFolderName <| childSourceDirectory subFolderName ma

let renderMarkdownFile (stylesheetName:string option)
                       (docTitle:string)
                       (markdown:MarkdownFile) : DocMonadWord<PdfFile> =
    docMonad {
        let! (stylesheet:WordFile option) = 
            match stylesheetName with
            | None -> dreturn None
            | Some name -> 
                askIncludeFile name >>= getWordFile >>= (dreturn << Some)
 
        let! docx = Markdown.markdownToWord stylesheet markdown
        let! pdf = WordFile.exportPdf PqScreen docx |>> setTitle docTitle
        return pdf
    }

let coversheet (siteName:string) (saiNumber:string) : DocMonadWord<PdfFile> = 
    docMonad { 
        let! logoPath = askIncludeFile "YW-logo.jpg"
        let! stylesPath = askIncludeFile "custom-reference1.docx"
        let! (stylesheet:WordFile option) = getWordFile stylesPath |>> Some
        let! markdownFile = coversheet saiNumber siteName logoPath "S Tetley" "coversheet.md" 
        let! docx = Markdown.markdownToWord stylesheet markdownFile 
        let! pdf = WordFile.exportPdf PqScreen docx |>> setTitle "Coversheet"
        return pdf
    }

type PhotosDocType = 
    | PhotosSurvey 
    | PhotosSiteWork

    member x.Name 
        with get() : string = 
            match x with
            | PhotosSurvey -> "Survey"
            | PhotosSiteWork -> "Site Work"

    member x.SourceSubPath 
        with get() : string = 
            match x with
            | PhotosSurvey -> "1.Survey" </> "PHOTOS"
            | PhotosSiteWork -> "2.Site_Work" </> "PHOTOS"

    member x.WorkingSubPath 
        with get() : string = 
            match x with
            | PhotosSurvey -> "Survey_Photos"
            | PhotosSiteWork -> "Site_Work_Photos"



let photosDoc  (docType:PhotosDocType) : DocMonadWord<PdfFile> = 
    docMonad { 
        let title = sprintf "%s Photos" docType.Name
        let outputFile = sprintf "%s Photos.md" docType.Name |> safeName
        let! md = makePhotoBook title docType.SourceSubPath docType.WorkingSubPath outputFile
        let! pdf = renderMarkdownFile (Some docxCustomReference) docType.WorkingSubPath md
        return pdf
    }


// May have multiple surveys...
let surveys () : DocMonadWord<PdfFile list> = 
    docMonad {
        let! inputs = findAllSourceFilesMatching "*Survey.doc*"
        let! pdfs = forM inputs (fun file -> getWordFile file >>= WordFile.exportPdf PqScreen)
        return pdfs
    }

let surveyPhotos () : DocMonadWord<PdfFile> = 
    photosDoc PhotosSurvey


let siteWorksPhotos () : DocMonadWord<PdfFile> = 
    photosDoc PhotosSiteWork

/// May have multiple documents
/// Get all doc files
let siteWorks () : DocMonadWord<PdfFile list> = 
    docMonad {
        let! inputs = findAllSourceFilesMatching "*.doc*"
        let! pdfs = forM inputs (fun file -> getWordFile file >>= WordFile.exportPdf PqScreen)
        return pdfs
    }





let getWorkList () : string list = 
    System.IO.DirectoryInfo(inputRoot).GetDirectories()
        |> Array.map (fun di -> di.Name)
        |> Array.toList

let buildOne (sourceName:string) : DocMonadWord<unit> = 
    localSubDirectory sourceName <| 
        docMonad {
            let siteName = getSiteName sourceName
            let! saiNumber = getSaiNumber siteName
            let! cover =  coversheet siteName saiNumber
            return ()
        }

let buildAll () : DocMonadWord<unit> = 
    let worklist = getWorkList ()
    foriMz worklist 
        <| fun ix name ->
                ignore <| printfn "Site %i: %s" (ix+1) name
                buildOne name

        


let demo01 () = 
    let userRes = new WordFile.WordHandle()
    runDocMonad userRes WindowsEnv 
        <| buildAll ()


let demo02 () = 
    let userRes = new WordFile.WordHandle()
    runDocMonad userRes WindowsEnv 
        <| childSourceDirectory @"AISLABY_CSO\1.Survey" (surveys ())
            

let demo03 () = 
    let userRes = new WordFile.WordHandle()
    runDocMonad userRes WindowsEnv 
        <| childSourceDirectory @"ABERFORD ROAD_NO 1 CSO\2.Site_work" (siteWorks ())

let demo04 () = 
    let userRes = new WordFile.WordHandle()
    runDocMonad userRes WindowsEnv 
        <| commonSubFolder @"ABERFORD ROAD_NO 1 CSO" (surveyPhotos ())

