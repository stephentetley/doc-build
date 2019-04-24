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

// ImageMagick
#I @"C:\Users\stephen\.nuget\packages\Magick.NET-Q8-AnyCPU\7.11.1\lib\netstandard20"
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
#I @"C:\Users\stephen\.nuget\packages\slformat\1.0.2-alpha-20190313\lib\netstandard2.0"
#r @"SLFormat.dll"
#I @"C:\Users\stephen\.nuget\packages\markdowndoc\1.0.1-alpha-20190314\lib\netstandard2.0"
#r @"MarkdownDoc.dll"


#load "..\src\DocBuild\Base\Internal\FakeLikePrim.fs"
#load "..\src\DocBuild\Base\Internal\FilePaths.fs"
#load "..\src\DocBuild\Base\Internal\Shell.fs"
#load "..\src\DocBuild\Base\Internal\GhostscriptPrim.fs"
#load "..\src\DocBuild\Base\Internal\PandocPrim.fs"
#load "..\src\DocBuild\Base\Internal\PdftkPrim.fs"
#load "..\src\DocBuild\Base\Internal\ImageMagickPrim.fs"
#load "..\src\DocBuild\Base\Common.fs"
#load "..\src\DocBuild\Base\DocMonad.fs"
#load "..\src\DocBuild\Base\Document.fs"
#load "..\src\DocBuild\Base\Collection.fs"
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
open DocBuild.Office.PandocWordShim

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




//let inputRoot   = @"G:\work\Projects\events2\final-docs\input\CSO_SPS"
//let outputRoot  = @"G:\work\Projects\events2\final-docs\output\CSO_SPS"

let (docxCustomReference:string) = @"custom-reference1.docx"

type DocMonadWord<'a> = DocMonad<WordDocument.WordHandle,'a>

let WindowsEnv : DocBuildEnv = 
    { WorkingDirectory = @"G:\work\Projects\events2\final-docs\output\CSO_SPS"
      SourceDirectory =  @"G:\work\Projects\events2\final-docs\input\CSO_SPS"
      IncludeDirectories = [ @"G:\work\Projects\events2\final-docs\input\include" ]
      PandocOpts = 
        { CustomStylesDocx = Some "custom-reference1.docx"
          PdfEngine = Some "pdflatex"
        }
      PrintOrScreen = PrintQuality.Screen
      }

let WindowsWordResources () : AppResources<WordDocument.WordHandle> = 
    let userRes = new WordDocument.WordHandle()
    { GhostscriptExe = @"C:\programs\gs\gs9.15\bin\gswin64c.exe"
      PdftkExe = @"pdftk"
      PandocExe = @"pandoc"
      UserResources = userRes
    }


let getSiteName (sourceName:string) : string = 
    sourceName.Replace("_", "/")




let renderMarkdownFile  (docTitle:string)
                       (markdown:MarkdownDoc) : DocMonadWord<PdfDoc> =
    docMonad {
        let! docx = Markdown.markdownToWord  markdown
        return! WordDocument.exportPdf docx |>> setTitle docTitle
    }

let genCoversheet (siteName:string) (saiNumber:string) : DocMonadWord<PdfDoc> = 
    docMonad { 
        let! logoPath = getIncludeJpegDoc "YW-logo.jpg"
        let config:CoversheetConfig = 
            { LogoPath = logoPath.AbsolutePath
              SaiNumber = saiNumber
              SiteName = siteName  
              Author = "S Tetley"
              Title = "T0975 - Event Duration Monitoring Phase 2 (EDM2)"
            }
        let! markdownFile = coversheet config "coversheet.md" 
        let! docx = Markdown.markdownToWord markdownFile 
        return! WordDocument.exportPdf docx |>> setTitle "Coversheet"
    }


let surveyPhotosConfig (siteName:string) : PhotoBookConfig = 
    { Title = sprintf "%s Survey Photos" siteName
      SourceSubFolder = "1.Survey" </> "PHOTOS"
      WorkingSubFolder = "Survey_Photos"
      RelativeOutputName = "survey_photos.md"
    }
              
let siteWorksPhotosConfig (siteName:string) : PhotoBookConfig = 
    { Title = sprintf "%s Site Work Photos" siteName
      SourceSubFolder = "2.Site_Work" </> "PHOTOS"
      WorkingSubFolder =  "Site_Work_Photos"
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
        return! WordDocument.exportPdf doc |>> setTitle title
        }

// May have multiple surveys...
let processSurveys (siteName:string) : DocMonadWord<PdfDoc list> = 
    docMonad {
        let! inputs = 
            localSourceSubdirectory "1.Survey" 
                <| findAllSourceFilesMatching "*Survey*.doc*" false
        return! mapM (wordDocToPdf siteName) inputs
        
    }

let genSurveyPhotos (siteName:string) : DocMonadWord<PdfDoc option> = 
    surveyPhotosConfig siteName 
        |> makePhotoBook
        |>> Option.map (setTitle "Survey Photos")

let genSiteWorkPhotos (siteName:string) : DocMonadWord<PdfDoc option> = 
    siteWorksPhotosConfig siteName 
        |> makePhotoBook
        |>> Option.map (setTitle "Site Work Photos")


let genContents (pdfs:PdfCollection) : DocMonadWord<PdfDoc> =
    let config : ContentsConfig = 
        { PrologLength = 1
          RelativeOutputName = "contents.md" }
    makeTableOfContents config pdfs

/// May have multiple documents
/// Get doc files matching glob 
// (run twice for Calibrations and RTU installs)
let processSiteWork (siteName:string) (glob:string) : DocMonadWord<PdfDoc list> = 
    docMonad {
        let! inputs = 
            localSourceSubdirectory "2.Site_work" 
                <| findAllSourceFilesMatching glob false
        return! mapM (wordDocToPdf siteName) inputs
    }

let processRTUInstalls (siteName:string) : DocMonadWord<PdfDoc list> = 
    processSiteWork siteName "*RTU Install*.doc*"

let processUSCalibrations (siteName:string) : DocMonadWord<PdfDoc list> = 
    processSiteWork siteName "*US Calib*.doc*"



let build1 (saiMap:SaiMap) : DocMonadWord<PdfDoc> = 
    docMonad {
        let! sourceName = askSourceDirectory () |>> fileObjectName
        let  siteName = getSiteName sourceName
        let! saiNumber = getSaiNumber saiMap siteName |> liftOption "No SAI Number"
        let! cover = genCoversheet siteName saiNumber
        let! surveys = processSurveys siteName
        let! oSurveyPhotos = genSurveyPhotos siteName
        let! calibrations = processUSCalibrations siteName
        let! rtuInstalls = processRTUInstalls siteName
        let! oWorksPhotos = genSiteWorkPhotos siteName
        let col1 = Collection.empty  
                        &^^ surveys 
                        &^^ oSurveyPhotos 
                        &^^ calibrations 
                        &^^ rtuInstalls
                        &^^ oWorksPhotos
        let! contents = genContents col1
        let colAll = cover ^^& contents ^^& col1
        let! outputAbsPath = extendWorkingPath (sprintf "%s Final.pdf" sourceName)
        return! Pdf.concatPdfs Pdf.GsDefault colAll outputAbsPath 
    }





let main () = 
    let resources = WindowsWordResources ()
    let saiMap = buildSaiMap ()
    runDocMonad resources WindowsEnv 
        <| foreachSourceIndividualOutput defaultSkeletonOptions 
                                         (foreachSourceIndividualOutput defaultSkeletonOptions (build1 saiMap))
