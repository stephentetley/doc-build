// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


#r "netstandard"
open System
open System.Text

// Office deps
#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.Word"
#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Excel\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.Excel"
#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.PowerPoint\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.PowerPoint"
#I @"C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c"
#r "office"
open Microsoft.Office.Interop

// ImageMagick
#I @"C:\Users\stephen\.nuget\packages\magick.net-q8-anycpu\7.11.1\lib\netstandard20"
#r @"Magick.NET-Q8-AnyCPU.dll"
#I @"C:\Users\stephen\.nuget\packages\magick.net-q8-anycpu\7.11.1\runtimes\win-x64\native"

// ExcelProvider
#I @"C:\Users\stephen\.nuget\packages\ExcelProvider\1.0.1\lib\netstandard2.0"
#r "ExcelProvider.Runtime.dll"

#I @"C:\Users\stephen\.nuget\packages\ExcelProvider\1.0.1\typeproviders\fsharp41\netstandard2.0"
#r "ExcelDataReader.DataSet.dll"
#r "ExcelDataReader.dll"
#r "ExcelProvider.DesignTime.dll"
open FSharp.Interop.Excel


// SLFormat & MarkdownDoc (not on nuget.org)
#I @"C:\Users\stephen\.nuget\packages\slformat\1.0.2-alpha-20190322\lib\netstandard2.0"
#r @"SLFormat.dll"
#I @"C:\Users\stephen\.nuget\packages\markdowndoc\1.0.1-alpha-20190508\lib\netstandard2.0"
#r @"MarkdownDoc.dll"


#load "..\src\DocBuild\Base\Internal\FilePaths.fs"
#load "..\src\DocBuild\Base\Internal\GhostscriptPrim.fs"
#load "..\src\DocBuild\Base\Internal\PandocPrim.fs"
#load "..\src\DocBuild\Base\Internal\PdftkPrim.fs"
#load "..\src\DocBuild\Base\Internal\ImageMagickPrim.fs"
#load "..\src\DocBuild\Base\Common.fs"
#load "..\src\DocBuild\Base\DocMonad.fs"
#load "..\src\DocBuild\Base\Document.fs"
#load "..\src\DocBuild\Base\Collection.fs"
#load "..\src\DocBuild\Base\FindFiles.fs"
#load "..\src\DocBuild\Base\FileOperations.fs"
#load "..\src\DocBuild\Base\Skeletons.fs"
#load "..\src\DocBuild\Document\Pdf.fs"
#load "..\src\DocBuild\Document\Jpeg.fs"
#load "..\src\DocBuild\Document\Markdown.fs"
#load "..\src\DocBuild\Extra\Contents.fs"
#load "..\src\DocBuild\Extra\PhotoBook.fs"
#load "..\src\DocBuild\Extra\TitlePage.fs"

#load "..\src-msoffice\DocBuild\Office\Internal\Utils.fs"
#load "..\src-msoffice\DocBuild\Office\Internal\WordPrim.fs"
#load "..\src-msoffice\DocBuild\Office\Internal\ExcelPrim.fs"
#load "..\src-msoffice\DocBuild\Office\Internal\PowerPointPrim.fs"
#load "..\src-msoffice\DocBuild\Office\WordDocument.fs"
#load "..\src-msoffice\DocBuild\Office\ExcelDocument.fs"
#load "..\src-msoffice\DocBuild\Office\PowerPointDocument.fs"
#load "..\src-msoffice\DocBuild\Office\PandocWordShim.fs"

open DocBuild.Base
open DocBuild.Document
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





type DocMonadWord<'a> = DocMonad<'a, WordDocument.WordHandle>

let WindowsEnv : DocBuildEnv = 
    { SourceDirectory       = @"G:\work\Projects\usar\final-docs\NSWC_mop_up_2\input"
      WorkingDirectory      = @"G:\work\Projects\usar\final-docs\NSWC_mop_up_2\output"
      IncludeDirectories    = [ @"G:\work\Projects\usar\final-docs\include" ]
      PrintOrScreen = PrintQuality.Screen
      PandocOpts = 
        { CustomStylesDocx = Some "custom-reference1.docx"
          PdfEngine = Some "pdflatex"
        }
    }


let WindowsWordResources () : AppResources<WordDocument.WordHandle> = 
    let userRes = new WordDocument.WordHandle()
    { GhostscriptExe = @"C:\programs\gs\gs9.15\bin\gswin64c.exe"
      PdftkExe = @"pdftk"
      PandocExe = @"pandoc"
      UserResources = userRes
    }


let getSiteName (folderName:string) : string = 
    folderName.Replace('_', '/')




let renderMarkdownFile (docTitle:string)
                       (markdown:MarkdownDoc) : DocMonadWord<PdfDoc> =
    docMonad {
        let! docx = Markdown.markdownToWord markdown
        return! WordDocument.exportPdf docx |>> setTitle docTitle
    }

let genCoversheet (siteName:string) (saiNumber:string option) : DocMonadWord<PdfDoc> = 
    docMonad { 
        let! logoPath = getIncludeJpegDoc "YW-logo.jpg"
        let config:CoversheetConfig = 
            { LogoPath = logoPath.AbsolutePath
              SaiNumber = saiNumber
              SiteName = siteName  
              Author = "S Tetley"
              Title = "AMP6 Ultrasonic Asset Replacement (Scheme code S3820)"
            }
        let! markdownFile = coversheet config "coversheet.md" 
        let! docx = Markdown.markdownToWord markdownFile 
        return! WordDocument.exportPdf docx |>> setTitle "Coversheet"
    }

    
let genContents (prologLength:int) (pdfs:PdfCollection) : DocMonadWord<PdfDoc> =
    let config : PandocWordShim.ContentsConfig = 
        { PrologLength = prologLength
          RelativeOutputName = "contents.md" }
    PandocWordShim.makeTableOfContents config pdfs

let genProjectScope () : DocMonadWord<PdfDoc> =  
    docMonad { 
        let! (input:PdfDoc) = getIncludePdfDoc "project-scope.pdf"
        return! copyDocumentToWorking input |>> setTitle "Project Scope"
    }





let surveyPhotosConfig (siteName:string) : PandocWordShim.PhotoBookConfig = 
    { Title = sprintf "%s Survey Photos" siteName
      SourceSubdirectory = "1.Survey" </> "PHOTOS"
      WorkingSubdirectory = "Survey_Photos"
      RelativeOutputName = "survey_photos.md"
    }
              
let siteWorksPhotosConfig (siteName:string) : PandocWordShim.PhotoBookConfig = 
    { Title = sprintf "%s Install Photos" siteName
      SourceSubdirectory = "2.Site_Work" </> "PHOTOS"
      WorkingSubdirectory =  "Site_Work_Photos"
      RelativeOutputName = "site_works_photos.md"
    }



let sourceFileToTitle (siteName:string) (filePath:string) : string = 
    let fileNameExt = IO.Path.GetFileName filePath
    let fileName = IO.Path.GetFileNameWithoutExtension fileNameExt
    let safe = safeName siteName
    let patt = sprintf "^%s (?<good>.*)" safe
    let rmatch = RegularExpressions.Regex.Match(fileName, patt)
    if rmatch.Success then        
        rmatch.Groups.["good"].Value
    else
        fileName


let wordDocToPdf (siteName:string) (absPath:string) : DocMonadWord<PdfDoc> = 
    let title = sourceFileToTitle siteName absPath
    docMonad { 
        let! doc = getSourceWordDoc absPath
        let! pdf1 = WordDocument.exportPdf doc 
        return! PandocWordShim.prefixWithTitlePage title None pdf1 |>> setTitle title
    }

let processMarkdown (title:string)
                    (sourceSubfolder: string option)
                    (glob:string) : DocMonadWord<PdfDoc list> = 
    docMonad {
        let! inputs = 
            match sourceSubfolder with
            | Some subdir -> localSourceSubdirectory subdir 
                                        (findSourceFilesMatching glob false) 
            | None -> findSourceFilesMatching glob false
        return! mapM (fun path -> 
                        getSourceMarkdownDoc path >>= fun md1 ->
                        copyDocumentToWorking md1 >>= fun md2 ->
                        renderMarkdownFile title md2) inputs
    }


// May have multiple surveys...
// Can fail - should be caught
let processSurveySheets (siteName:string) : DocMonadWord<PdfDoc list> = 
    docMonad {
        let! inputs = 
            localSourceSubdirectory "1.Survey" 
                <| (assertNonEmpty =<< findSourceFilesMatching "*Survey*.doc*" false)
        return! mapM (wordDocToPdf siteName) inputs
    }

let processSurveys (siteName:string) : DocMonadWord<PdfDoc list> = 
    processSurveySheets siteName 
        <|> processMarkdown "Survey Info" (Some "1.Survey") "*.md"

let genSurveyPhotos (siteName:string) : DocMonadWord<PdfDoc> = 
    PandocWordShim.makePhotoBook (surveyPhotosConfig siteName) 
        |>> setTitle "Survey Photos"

let genSiteWorkPhotos (siteName:string) : DocMonadWord<PdfDoc> = 
    PandocWordShim.makePhotoBook (siteWorksPhotosConfig siteName) 
        |>> setTitle "Site Work Photos"



/// May have multiple documents
/// Get doc files matching glob 
let processInstallSheets (siteName:string) : DocMonadWord<PdfDoc list> = 
    docMonad {
        let! inputs = 
            localSourceSubdirectory "2.Site_work"
                    <| (assertNonEmpty =<< findSourceFilesMatching "*Install*.doc*" false)
        return! mapM (wordDocToPdf siteName) inputs
    }

let processInstalls(siteName:string) : DocMonadWord<PdfDoc list> = 
    processInstallSheets siteName 
        <|> processMarkdown "Installation Info" (Some "2.Site_work") "*.md"


let build1 (saiMap:SaiMap) : DocMonadWord<PdfDoc> = 
    docMonad {
        let! sourceName = askSourceDirectory () |>> fileObjectName
        let siteName = getSiteName sourceName
        let saiNumber = getSaiNumber saiMap siteName
        let! cover = genCoversheet siteName saiNumber
        let! scope = genProjectScope ()
        let! surveys = processSurveys siteName
        let! surveyPhotos = genSurveyPhotos siteName
        let! ultras = processInstalls siteName
        let! worksPhotos = genSiteWorkPhotos siteName
        let (col1 : PdfCollection) =  
            Collection.concat
                 [ Collection.singleton scope
                 ; Collection.ofList surveys 
                 ; Collection.singleton surveyPhotos 
                 ; Collection.ofList ultras
                 ; Collection.singleton worksPhotos
                 ]
        let! prologLength = Pdf.countPages cover
        let! contents = genContents prologLength col1
        let allDocs = cover ^^ contents ^^ col1
        let finalName = sprintf "%s Final.pdf" sourceName |> safeName
        return! Pdf.concatPdfs Pdf.GsDefault finalName allDocs 
    }


let main () = 
    let saiMap = buildSaiMap ()
    let resources = WindowsWordResources ()
    (resources.UserResources :> WordDocument.IWordHandle).PaperSizeForWord <- Some Word.WdPaperSize.wdPaperA4
    runDocMonad resources WindowsEnv 
        <| foreachSourceDirectory defaultSkeletonOptions 
                (build1 saiMap)

