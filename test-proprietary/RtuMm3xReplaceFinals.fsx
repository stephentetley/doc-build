// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


#r "netstandard"
open System
open System.Text.RegularExpressions

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
#I @"C:\Users\stephen\.nuget\packages\slformat\1.0.2-alpha-20190721\lib\netstandard2.0"
#r @"SLFormat.dll"
#I @"C:\Users\stephen\.nuget\packages\markdowndoc\1.0.1-alpha-20191014\lib\netstandard2.0"
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
open DocBuild.Office.PandocWordShim

#load "ExcelProviderHelper.fs"
#load "Proprietary.fs"
open ExcelProviderHelper
open Proprietary

// ImageMagick Dll loader.
// A hack to get over Dll loading error due to the 
// native dll `Magick.NET-Q8-x64.Native.dll`
[<Literal>] 
let NativeMagick = @"C:\Users\stephen\.nuget\packages\magick.net-q8-anycpu\7.9.2\runtimes\win-x64\native"
Environment.SetEnvironmentVariable("PATH", 
    Environment.GetEnvironmentVariable("PATH") + ";" + NativeMagick
    )



let WindowsEnv : DocBuildEnv = 
    { SourceDirectory   = @"G:\work\Projects\rtu\final-docs\input\y5-mm3x-replacements-batch1"
      WorkingDirectory  = @"G:\work\Projects\rtu\final-docs\output\y5-mm3x-replacements-batch1"
      IncludeDirectories = [ @"G:\work\Projects\rtu\final-docs\include" ]
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

type DocMonadWord<'a> = DocMonad<'a, WordDocument.WordHandle>

let getSiteName (sourceName:string) : string = 
    sourceName.Replace("_", "/")

let renderMarkdownDoc (docTitle:string)
                      (markdown:MarkdownDoc) : DocMonadWord<PdfDoc> =
    docMonad {
        let! docx = Markdown.markdownToWord markdown
        return! WordDocument.exportPdf  docx |>> setTitle docTitle
    }


let coverSeaches (saiNumber: String) (siteName: string) (year: string) : SearchList = 
    [ ("#SAINUM", saiNumber)
    ; ("#SITENAME", siteName) 
    ; ("#YEAR", year)
    ; ("#DATE", DateTime.Today.ToString(format = "dd/MM/yyyy") )
    ]


let genCover (saiNumber: String) (siteName: string) (year: string)  : DocMonadWord<PdfDoc> = 
    let outputName = sprintf "%s cover.docx" (siteName |> safeName )
    let searches : SearchList = coverSeaches saiNumber siteName year
    docMonad { 
        let! (template:WordDoc) = getIncludeWordDoc "TEMPLATE MM3x-to-MMIM Cover Sheet.docx"
        let! outpath = extendWorkingPath outputName
        let! wordFile = WordDocument.findReplaceAs searches outpath template
        return! WordDocument.exportPdf wordFile
    }


/// Folder1 either "1.Survey" or "2.Site_work"
let sourceWordDocToPdf (folder1:string) (fileGlob:string) : DocMonadWord<PdfDoc> = 
    let updateRes handle = 
        (handle :> WordDocument.IWordHandle).PaperSizeForWord <- None
        handle
    
    localUserResources (updateRes)
        << localSourceSubdirectory (folder1) 
        <| docMonad { 
                let! input = assertExactlyOne =<< findSourceFilesMatching fileGlob false 
                let! doc = getWordDoc input
                return! WordDocument.exportPdf doc
            }

let processMarkdown1 (title : string)
                     (sourceSubfolder : string)
                     (glob : string) : DocMonadWord<PdfDoc> = 
    docMonad {
        let! input = 
            localSourceSubdirectory sourceSubfolder
                <| (assertExactlyOne =<< findSourceFilesMatching glob false)
        let! md = getSourceMarkdownDoc input
        return! renderMarkdownDoc title md
    }

let genSurvey () :DocMonadWord<PdfDoc> = 
    sourceWordDocToPdf "1.Survey" "*urvey*.doc*"
        <|> processMarkdown1 "Survey" "1.Survey" "*.md"

let genSiteWorks () :DocMonadWord<PdfDoc> = 
    sourceWordDocToPdf "2.Site_work" "*Site*Works*.doc*"
        <|> processMarkdown1 "Site Work" "2.Site_work" "*.md"
                

let genSurveyPhotos (siteName: String) : DocMonadWord<PdfDoc> = 
    let name1 = safeName siteName
    let props : PandocWordShim.PhotoBookConfig = 
        { Title = "Survey Photos"
        ; SourceSubdirectory = "1.Survey" </> "photos"
        ; WorkingSubdirectory = "survey_photos"
        ; RelativeOutputName = sprintf "%s survey photos.md" name1 }
    PandocWordShim.makePhotoBook props 


let genSiteWorkPhotos (siteName: String) : DocMonadWord<PdfDoc> = 
    let name1 = safeName siteName
    let props : PandocWordShim.PhotoBookConfig = 
        { Title = "Install Photos"
        ; SourceSubdirectory  = "2.Site_work" </> "photos"
        ; WorkingSubdirectory = "install_photos"
        ; RelativeOutputName= sprintf "%s install photos.md" name1 }
    PandocWordShim.makePhotoBook props 
    

let build1 (saiMap:SaiMap) (year: string) : DocMonadWord<PdfDoc> = 
    docMonad { 
        let! name1 = sourceDirectoryName ()
        let  siteName = getSiteName name1
        let! saiNumber = liftOption "sai not found" <| getSaiNumber saiMap siteName

        let! cover = genCover saiNumber siteName year
        let! survey = mandatory "Survey" <| genSurvey ()
        let! surveyPhotos = nonMandatory <| genSurveyPhotos siteName
        let! siteWorks = mandatory "Site Work" <| genSiteWorks ()
        let! worksPhotos = nonMandatory <| genSiteWorkPhotos siteName

        let (col1:PdfCollection) = 
            Collection.ofList [ cover; survey; surveyPhotos; siteWorks; worksPhotos]

        let finalName = sprintf "%s Final.pdf" siteName |> safeName
        return! Pdf.concatPdfs Pdf.GsDefault finalName col1
    }

// work type = "MM3X to MK5 MMIM Asset Replacement"

let main () = 
    let resources = WindowsWordResources ()
    let saiMap = buildSaiMap ()
    let yearName = "Year 5"
    let options = defaultSkeletonOptions // { defaultSkeletonOptions with TestingSample = TakeDirectories 5 }
    runDocMonad resources WindowsEnv 
        <| foreachSourceDirectory options (build1 saiMap yearName)

