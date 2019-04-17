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
    let includePath = DirectoryPath @"G:\work\Projects\rtu\final-docs\include"
    { WorkingDirectory = DirectoryPath @"G:\work\Projects\rtu\IS_barriers\final-docs\output\Batch04"
      SourceDirectory =  DirectoryPath @"G:\work\Projects\rtu\IS_barriers\final-docs\input\batch4_finals_source"
      IncludeDirectory = includePath
      GhostscriptExe = @"C:\programs\gs\gs9.15\bin\gswin64c.exe"
      PdftkExe = @"pdftk"
      PandocExe = @"pandoc" 
      PrintOrScreen = PrintQuality.Screen
      CustomStylesDocx = Some (includePath <//> @"custom-reference1.docx")
      PandocPdfEngine = Some "pdflatex"
      }

type DocMonadWord<'a> = DocMonad<WordDocument.WordHandle,'a>



let renderMarkdownDoc (docTitle:string)
                      (markdown:MarkdownDoc) : DocMonadWord<PdfDoc> =
    docMonad {
        let! docx = Markdown.markdownToWord markdown
        return! WordDocument.exportPdf  docx |>> setTitle docTitle
    }



let sourceWordDocToPdf (fileGlob:string) :DocMonadWord<PdfDoc option> = 
    docMonad { 
        let! input = tryFindExactlyOneSourceFileMatching fileGlob false
        match input with
        | None -> return None
        | Some infile ->
            let! doc = getWordDoc infile
            return! (WordDocument.exportPdf doc |>> Some)
    }


    

let genSiteWorks () : DocMonadWord<PdfDoc> = 
    optionFailM "No Site Works document" 
                (sourceWordDocToPdf "*Site Works*.doc*")
                

//let genPhotos (siteName:string) : DocMonadWord<PdfDoc option> = 
//    let name1 = safeName row.``Site Name``
//    let props : PhotoBookConfig = 
//        { Title = "Survey Photos"
//        ; SourceSubFolder = "1.Surveys" </> name1 </> "photos"
//        ; WorkingSubFolder = "survey_photos"
//        ; RelativeOutputName = sprintf "%s survey photos.md" name1 }
//    makePhotoBook props

let genFinalDoc1 () :DocMonadWord<unit> = 
    docMonad { 
        let! workSheet = genSiteWorks ()
        return ()
    }

//let genFinal (sourceFolderName:string) :DocMonadWord<PdfDoc> = 
//    localWorkingSubdirectory (sourceFolderName) 
//        <| docMonad { 
//                let! cover = genCover row 
//                let! oSurvey = genSurvey row
//                let! works = genSiteWorks row
//                let! oSurveyPhotos = genSurveyPhotos row
//                let! oWorksPhotos = genWorkPhotos row

//                let (col:PdfCollection) = 
//                    Collection.singleton cover 
//                        &^^ oSurvey     &^^ oSurveyPhotos
//                        &^^ works       &^^ oWorksPhotos

//                let! outputAbsPath = extendWorkingPath (sprintf "%s Final.pdf" safeSiteName)
//                return! Pdf.concatPdfs Pdf.GsDefault col outputAbsPath 
//            }


let isLike (pattern:string) (source:string) = 
    Regex.IsMatch(input=source, pattern=pattern)


let getWorkList () : string list = 
    System.IO.DirectoryInfo(WindowsEnv.SourceDirectory.LocalPath).GetDirectories()
        |> Array.map (fun info -> info.Name)
        |> Array.toList

let getSourceChildren () : DocMonad<'res,string list> = 
    docMonad { 
        let! source = askSourceDirectoryPath ()
        let! (kids: IO.DirectoryInfo[]) = liftIO <| fun _ -> System.IO.DirectoryInfo(source).GetDirectories() 
        return (kids |> Array.map (fun info -> info.Name) |> Array.toList)
    }



let forallSourceChildren (process1: DocMonad<'res,'a>) : DocMonad<'res, unit> = 
    getSourceChildren () >>= fun srcDirs -> 
    forMz srcDirs (fun dir -> commonSubdirectory dir process1)

let main () = 
    let userRes = new WordDocument.WordHandle()
    runDocMonad userRes WindowsEnv 
        <| forallSourceChildren (genFinalDoc1 ())