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
#I @"C:\Users\stephen\.nuget\packages\slformat\1.0.2-alpha-20190313\lib\netstandard2.0"
#r @"SLFormat.dll"
#I @"C:\Users\stephen\.nuget\packages\markdowndoc\1.0.1-alpha-20190314\lib\netstandard2.0"
#r @"MarkdownDoc.dll"


#load "..\src\DocBuild\Base\BaseDefinitions.fs"
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
open ExcelProviderHelper

// ImageMagick Dll loader.
// A hack to get over Dll loading error due to the 
// native dll `Magick.NET-Q8-x64.Native.dll`
[<Literal>] 
let NativeMagick = @"C:\Users\stephen\.nuget\packages\magick.net-q8-anycpu\7.9.2\runtimes\win-x64\native"
Environment.SetEnvironmentVariable("PATH", 
    Environment.GetEnvironmentVariable("PATH") + ";" + NativeMagick
    )



let WindowsEnv : DocBuildEnv = 
    { WorkingDirectory = DirectoryPath @"G:\work\Projects\rtu\IS_barriers\final-docs\output\Batch04"
      SourceDirectory =  DirectoryPath @"G:\work\Projects\rtu\IS_barriers\final-docs\input\batch4_finals_source"
      IncludeDirectory = DirectoryPath @"G:\work\Projects\rtu\final-docs\include"
      PrintOrScreen = PrintQuality.Screen
      PandocOpts = 
        {  
          CustomStylesDocx = Some "custom-reference1.docx"
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

type DocMonadWord<'a> = DocMonad<WordDocument.WordHandle,'a>

let sourceToSiteName (sourceName:string) : string = 
    sourceName.Replace("_", "/")


let sourceWordDocToPdf (fileGlob:string) : DocMonadWord<PdfDoc option> = 
    docMonad { 
        match! tryFindExactlyOneSourceFileMatching fileGlob false with
        | None -> return None
        | Some infile ->
            let! doc = getWordDoc infile
            return! (WordDocument.exportPdf doc |>> Some)
    }


    

let genSiteWorks () : DocMonadWord<PdfDoc> = 
    optionFailM "No Site Works document" 
                (sourceWordDocToPdf "*Site Works*.doc*")
                

let genPhotos (siteName:string) : DocMonadWord<PdfDoc option> = 
    let props : PhotoBookConfig = 
        { Title = "Site Photos"
        ; SourceSubFolder = "photos"
        ; WorkingSubFolder = "photos"
        ; RelativeOutputName = sprintf "%s survey photos.md" (safeName siteName) }
    makePhotoBook props

let genFinalDoc1 () : DocMonadWord<PdfDoc> = 
    docMonad { 
        let! sourceName =  askSourceDirectoryName ()
        let siteName = sourceName |> sourceToSiteName
        let! workSheet = genSiteWorks ()
        let! phodoDoc = genPhotos siteName
        let (col:PdfCollection) = 
            Collection.empty 
                &^^ workSheet     &^^ phodoDoc

        let! outputAbsPath = extendWorkingPath (sprintf "%s S3953 IS Barrier Final.pdf" sourceName)
        return! Pdf.concatPdfs Pdf.GsDefault col outputAbsPath 
    }



let main () = 
    let res = WindowsWordResources ()
    runDocMonad res WindowsEnv 
        <| dtodSourceChildren defaultSkeletonOptions (genFinalDoc1 ())