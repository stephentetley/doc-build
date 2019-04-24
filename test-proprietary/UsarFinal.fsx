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
#I @"C:\Users\stephen\.nuget\packages\slformat\1.0.2-alpha-20190313\lib\netstandard2.0"
#r @"SLFormat.dll"
#I @"C:\Users\stephen\.nuget\packages\markdowndoc\1.0.1-alpha-20190314\lib\netstandard2.0"
#r @"MarkdownDoc.dll"

#load "..\src\DocBuild\Base\FakeLikePrim.fs"
#load "..\src\DocBuild\Base\FilePaths.fs"
#load "..\src\DocBuild\Base\Common.fs"
#load "..\src\DocBuild\Base\Shell.fs"
#load "..\src\DocBuild\Base\DocMonad.fs"
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
open DocBuild.Base.DocMonad
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





type DocMonadWord<'a> = DocMonad<WordDocument.WordHandle,'a>

let WindowsEnv : DocBuildEnv = 
    { WorkingDirectory      = @"G:\work\Projects\usar\final-docs\small_stw_mopup\output"
      SourceDirectory       = @"G:\work\Projects\usar\final-docs\small_stw_mopup\input"
      IncludeDirectories    = [ @"G:\work\Projects\usar\final-docs\include" ]
      PrintOrScreen = PrintQuality.Screen
      PandocOpts = 
        { CustomStylesDocx = Some "custom-reference1.docx"
          PdfEngine = Some "pdflatex"
        }
    }


let WindowsWordResources () : Resources<WordDocument.WordHandle> = 
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

let genCoversheet (siteName:string) (saiNumber:string) : DocMonadWord<PdfDoc> = 
    docMonad { 
        let! logoPath = includeJpegDoc "YW-logo.jpg"
        let config:CoversheetConfig = 
            { LogoPath = logoPath.AbsolutePath
              SaiNumber = saiNumber
              SiteName = siteName  
              Author = "S Tetley"
              Title = "AMP6 Ultrasonic Asset Replacement (Scheme code S3820)"
            }
        let! (stylesheet:WordDoc option) = includeWordDoc "custom-reference1.docx" |>> Some
        let! markdownFile = coversheet config "coversheet.md" 
        let! docx = Markdown.markdownToWord markdownFile 
        return! WordDocument.exportPdf docx |>> setTitle "Coversheet"
    }

let genProjectScope () : DocMonadWord<PdfDoc> =  
    docMonad { 
        let! (input:PdfDoc) = includePdfDoc "project-scope.pdf"
        return! copyToWorking input |>> setTitle "Project Scope"
    }

let surveyPhotosConfig (siteName:string) : PhotoBookConfig = 
    { Title = sprintf "%s Survey Photos" siteName
      SourceSubFolder = "1.Survey" </> "PHOTOS"
      WorkingSubFolder = "Survey_Photos"
      RelativeOutputName = "survey_photos.md"
    }
              
let siteWorksPhotosConfig (siteName:string) : PhotoBookConfig = 
    { Title = sprintf "%s Install Photos" siteName
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
        let! doc = sourceWordDoc absPath
        let! pdf1 = WordDocument.exportPdf doc 
        return! prefixWithTitlePage title None pdf1 |>> setTitle title
    }

let processMarkdown (title:string)
                    (subfolder: string option)
                    (glob:string) : DocMonadWord<PdfDoc list> = 
    docMonad {
        let contextM ma = 
            match subfolder with 
            | None -> ma
            | Some name -> localSourceSubdirectory name ma
        let! inputs = contextM  <| findSomeSourceFilesMatching glob false
        return! mapM (fun path -> 
                        sourceMarkdownDoc path >>= fun md1 ->
                        copyToWorking md1 >>= fun md2 ->
                        renderMarkdownFile title md2) inputs
    }


// May have multiple surveys...
// Can fail - should be caught
let processSurveySheets (siteName:string) : DocMonadWord<PdfDoc list> = 
    docMonad {
        let! inputs = 
            localSourceSubdirectory "1.Survey" 
                <| findSomeSourceFilesMatching "*Survey*.doc*" false
        return! mapM (wordDocToPdf siteName) inputs
    }

let processSurveys (siteName:string) : DocMonadWord<PdfDoc list> = 
    processSurveySheets siteName 
        <||> processMarkdown "Survey Info" (Some "1.Survey") "*.md"

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
        { PrologLength = 2
          RelativeOutputName = "contents.md" }
    makeTableOfContents config pdfs

/// May have multiple documents
/// Get doc files matching glob 
let processInstallSheets (siteName:string) : DocMonadWord<PdfDoc list> = 
    docMonad {
        let! inputs = 
            localSourceSubdirectory "2.Site_work"
                    <| findSomeSourceFilesMatching "*Install*.doc*" false
        return! mapM (wordDocToPdf siteName) inputs
    }

let processInstalls(siteName:string) : DocMonadWord<PdfDoc list> = 
    processInstallSheets siteName 
        <||> processMarkdown "Installation Info" (Some "2.Site_work") "*.md"



let getWorkList () : string list = 
    System.IO.DirectoryInfo(WindowsEnv.SourceDirectory).GetDirectories()
        |> Array.map (fun info -> info.Name)
        |> Array.toList

let buildOne (sourceName:string) 
             (siteName:string) 
             (saiNumber:string) : DocMonadWord<PdfDoc> = 
    commonSubdirectory sourceName <| 
        docMonad {
            let! cover = genCoversheet siteName saiNumber
            let! scope = genProjectScope ()
            let! surveys = processSurveys siteName
            let! oSurveyPhotos = genSurveyPhotos siteName
            let! ultras = processInstalls siteName
            let! oWorksPhotos = genSiteWorkPhotos siteName
            let col1 = Collection.empty  
                            &^^ scope
                            &^^ surveys 
                            &^^ oSurveyPhotos 
                            &^^ ultras
                            &^^ oWorksPhotos
            let! contents = genContents col1
            let allDocs = cover ^^& contents ^^& col1
            let! outputAbsPath = extendWorkingPath (sprintf "%s Final.pdf" sourceName)
            return! Pdf.concatPdfs Pdf.GsDefault allDocs outputAbsPath 
        }


let build1 (saiMap:SaiMap) (sourceName:string) : DocMonadWord<PdfDoc option> = 
    let siteName = getSiteName sourceName
    printfn "Site name: %s" siteName
    match getSaiNumber saiMap siteName with
    | None -> printfn "No sai"; mreturn None
    | Some sai -> 
        fmapM Some <| buildOne sourceName siteName sai


let buildAll () : DocMonadWord<unit> = 
    let worklist = getWorkList () 
    let saiMap = buildSaiMap ()
    foriMz worklist 
        <| fun ix name ->
                printfn "Site %i: %s" (ix+1) name |> ignore
                build1 saiMap name |>> ignore


let main () = 
    let resources = WindowsWordResources ()
    (resources.UserResources :> WordDocument.IWordHandle).PaperSizeForWord <- Some Word.WdPaperSize.wdPaperA4
    runDocMonad resources WindowsEnv 
        <| buildAll ()

