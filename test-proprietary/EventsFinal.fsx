// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


#r "netstandard"
open System


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
#I @"C:\Users\stephen\.nuget\packages\magick.net-q8-anycpu\7.9.2\lib\netstandard20"
#r @"Magick.NET-Q8-AnyCPU.dll"
#I @"C:\Users\stephen\.nuget\packages\magick.net-q8-anycpu\7.9.2\runtimes\win-x64\native"

// ExcelProvider
#I @"C:\Users\stephen\.nuget\packages\ExcelProvider\1.0.1\lib\netstandard2.0"
#r "ExcelProvider.Runtime.dll"

#I @"C:\Users\stephen\.nuget\packages\ExcelProvider\1.0.1\typeproviders\fsharp41\netstandard2.0"
#r "ExcelDataReader.DataSet.dll"
#r "ExcelDataReader.dll"
#r "ExcelProvider.DesignTime.dll"
open FSharp.Interop.Excel


// SLFormat & MarkdownDoc (not on nuget.org)
#I @"C:\Users\stephen\.nuget\packages\slformat\1.0.2-alpha-20190207\lib\netstandard2.0"
#r @"SLFormat.dll"
#I @"C:\Users\stephen\.nuget\packages\markdowndoc\1.0.1-alpha-20190226d\lib\netstandard2.0"
#r @"MarkdownDoc.dll"



#load "..\src\DocBuild\Base\FakeLikePrim.fs"
#load "..\src\DocBuild\Base\FilePaths.fs"
#load "..\src\DocBuild\Base\Common.fs"
#load "..\src\DocBuild\Base\Shell.fs"
#load "..\src\DocBuild\Base\DocMonad.fs"
#load "..\src\DocBuild\Base\DocMonadOperators.fs"
#load "..\src\DocBuild\Base\Document.fs"
#load "..\src\DocBuild\Base\Collection.fs"
#load "..\src\DocBuild\Base\FileOperations.fs"
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
#load "..\src-msoffice\DocBuild\Office\WordDocument.fs"
#load "..\src-msoffice\DocBuild\Office\ExcelDocument.fs"
#load "..\src-msoffice\DocBuild\Office\PowerPointDocument.fs"

open DocBuild.Base
open DocBuild.Base.DocMonad
open DocBuild.Base.DocMonadOperators
open DocBuild.Document
open DocBuild.Extra.PhotoBook
open DocBuild.Office

#load "ExcelProviderHelper.fs"
#load "Proprietary.fs"
#load "Coversheet.fs"
open Proprietary
open Coversheet



// ImageMagick Dll loader.
// A hack to get over Dll loading error due to the 
// native dll `Magick.NET-Q8-x64.Native.dll`
[<Literal>] 
let NativeMagick = @"C:\Users\stephen\.nuget\packages\magick.net-q8-anycpu\7.9.2\runtimes\win-x64\native"
Environment.SetEnvironmentVariable("PATH", 
    Environment.GetEnvironmentVariable("PATH") + ";" + NativeMagick
    )




let inputRoot   = @"G:\work\Projects\events2\final-docs\input\CSO_SPS"
let outputRoot  = @"G:\work\Projects\events2\final-docs\output\CSO_SPS"

let (docxCustomReference:string) = @"custom-reference1.docx"

type DocMonadWord<'a> = DocMonad<WordDocument.WordHandle,'a>

let WindowsEnv : BuilderEnv = 
    { WorkingDirectory = DirectoryPath @"G:\work\Projects\events2\final-docs\output\CSO_SPS"
      SourceDirectory =  DirectoryPath @"G:\work\Projects\events2\final-docs\input\CSO_SPS"
      IncludeDirectory = DirectoryPath @"G:\work\Projects\events2\final-docs\input\include"
      GhostscriptExe = @"C:\programs\gs\gs9.15\bin\gswin64c.exe"
      PdftkExe = @"pdftk"
      PandocExe = @"pandoc" }


let getSiteName (folderName:string) : string = 
    folderName.Replace('_', '/')


let getSaiNumber (siteName:string) : DocMonad<'res,string> = 
    mreturn "SAI00001234"       // TEMP



let renderMarkdownFile (stylesheetName:string option)
                       (docTitle:string)
                       (markdown:MarkdownDoc) : DocMonadWord<PdfDoc> =
    docMonad {
        let! (stylesheet:WordDoc option) = 
            match stylesheetName with
            | None -> mreturn None
            | Some name -> includeWordDoc name |>> Some
 
        let! docx = Markdown.markdownToWord stylesheet markdown
        let! pdf = WordDocument.exportPdf PqScreen docx |>> setTitle docTitle
        return pdf
    }

let genCoversheet (siteName:string) (saiNumber:string) : DocMonadWord<PdfDoc> = 
    docMonad { 
        let! logoPath = extendIncludePath "YW-logo.jpg"
        let! (stylesheet:WordDoc option) = includeWordDoc "custom-reference1.docx" |>> Some
        let! markdownFile = coversheet saiNumber siteName logoPath "S Tetley" "coversheet.md" 
        let! docx = Markdown.markdownToWord stylesheet markdownFile 
        let! pdf = WordDocument.exportPdf PqScreen docx |>> setTitle "Coversheet"
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



let photosDoc  (docType:PhotosDocType) : DocMonadWord<PdfDoc> = 
    docMonad { 
        let title = sprintf "%s Photos" docType.Name
        let outputFile = sprintf "%s Photos.md" docType.Name |> safeName
        let! md = makePhotoBook title docType.SourceSubPath docType.WorkingSubPath outputFile
        let! pdf = renderMarkdownFile (Some docxCustomReference) docType.WorkingSubPath md
        return pdf
    }


// May have multiple surveys...
let processSurveys () : DocMonadWord<PdfDoc list> = 
    docMonad {
        let! inputs = 
            localSourceSubdirectory "1.Survey" 
                <| findAllSourceFilesMatching "*Survey*.doc*" false
        let! pdfs = forM inputs (sourceWordDoc >=> WordDocument.exportPdf PqScreen)
        return pdfs
    }

let genSurveyPhotos () : DocMonadWord<PdfDoc> = 
    photosDoc PhotosSurvey


let genSiteWorkPhotos () : DocMonadWord<PdfDoc> = 
    photosDoc PhotosSiteWork

/// May have multiple documents
/// Get all doc files
let processSiteWork () : DocMonadWord<PdfDoc list> = 
    docMonad {
        let! inputs = 
            localSourceSubdirectory "2.Site_work" 
                <| findAllSourceFilesMatching "*.doc*" false
        let! pdfs = forM inputs (sourceWordDoc >=> WordDocument.exportPdf PqScreen)
        return pdfs
    }





let getWorkList () : string list = 
    System.IO.DirectoryInfo(inputRoot).GetDirectories()
        |> Array.map (fun di -> di.Name)
        |> Array.toList

let buildOne (sourceName:string) 
             (siteName:string) 
             (saiNumber:string) : DocMonadWord<PdfDoc> = 
    commonSubdirectory sourceName <| 
        docMonad {
            let! cover = genCoversheet siteName saiNumber
            let! surveys = processSurveys ()
            let! surveyPhotos = optionalM <| genSurveyPhotos ()
            let! worksheets = processSiteWork ()
            let! worksPhotos = optionalM <| genSiteWorkPhotos ()
            let col = Collection.singleton cover 
                            &>> surveys &>> surveyPhotos &>> worksheets &>> worksPhotos
            let! outputAbsPath = extendWorkingPath (sprintf "%s Final.pdf" sourceName)
            return! pdfConcat GsQuality.GsScreen outputAbsPath col
        }

let buildAll () : DocMonadWord<unit> = 
    let worklist = getWorkList ()
    foriMz worklist 
        <| fun ix name ->
                ignore <| printfn "Site %i: %s" (ix+1) name
                buildOne name "TEMP" "SAI01"

        


let demo01 () = 
    let userRes = new WordDocument.WordHandle()
    runDocMonad userRes WindowsEnv 
        <| commonSubdirectory @"ABERFORD ROAD_NO 1 CSO" (genCoversheet @"ABERFORD ROAD/NO 1 CSO" "SAI00036945")


let demo02 () = 
    let userRes = new WordDocument.WordHandle()
    runDocMonad userRes WindowsEnv 
        <| commonSubdirectory @"ABERFORD ROAD_NO 1 CSO" (processSurveys ())
            

let demo03 () = 
    let userRes = new WordDocument.WordHandle()
    runDocMonad userRes WindowsEnv 
        <| commonSubdirectory @"ABERFORD ROAD_NO 1 CSO" (processSiteWork ())


let demo04 () = 
    let userRes = new WordDocument.WordHandle()
    runDocMonad userRes WindowsEnv 
        <| commonSubdirectory @"ABERFORD ROAD_NO 1 CSO" (genSurveyPhotos ())

let demo05 () = 
    let userRes = new WordDocument.WordHandle()
    runDocMonad userRes WindowsEnv 
        <| buildOne @"AGBRIGG GARAGE_CSO" @"AGBRIGG GARAGE/CSO" "SAI00017527"

let demo06 () = 
    readADBAll () 
