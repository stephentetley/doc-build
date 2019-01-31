// Copyright (c) Stephen Tetley 2018,2019
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


open System

#load "..\src\DocBuild\Base\FakeLikePrim.fs"
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
#load "..\src-msoffice\DocBuild\Office\WordFile.fs"
#load "..\src-msoffice\DocBuild\Office\ExcelFile.fs"
#load "..\src-msoffice\DocBuild\Office\PowerPointFile.fs"

open DocBuild.Base
open DocBuild.Document.Pdf
open DocBuild.Base.DocMonad
open DocBuild.Base.DocMonadOperators

open DocBuild.Office

let WindowsEnv : BuilderEnv = 
    let cwd = System.IO.Path.Combine(__SOURCE_DIRECTORY__, "..", "data") |> folderUri
    { WorkingDirectory = cwd
      SourceDirectory = cwd
      IncludeDirectory = cwd <//> "include\\"
      GhostscriptExe = @"C:\programs\gs\gs9.15\bin\gswin64c.exe"
      PdftkExe = @"pdftk"
      PandocExe = @"pandoc"
    }


let demo01 () = 
    runDocMonadNoCleanup () WindowsEnv <| 
        docMonad { 
            let! p1 = workingPdfFile "One.pdf"
            let! p2 = workingPdfFile "Two.pdf" 
            let! p3 = workingPdfFile "Three.pdf"
            let pdfs = Collection.fromList [p1;p2;p3]
            let! outfile = workingPdfFile "Concat.pdf"
            let! _ = pdfConcat GsScreen outfile.LocalPath pdfs
            return ()
        }


let demo02 () = 
    runDocMonadNoCleanup () WindowsEnv <| 
        docMonad { 
            let! p1 = workingPdfFile "Concat.pdf"
            let! pageCount = pdfPageCount p1
            return pageCount
        }



let demo03 () = 
    let userRes = new WordFile.WordHandle()
    runDocMonad userRes WindowsEnv <| 
        docMonad { 
            let! w1 = workingWordFile "sample.docx" 
            let! p1 = WordFile.exportPdf PqScreen w1 
            return p1
        }

let demo04 () = 
    let userRes = new WordFile.WordHandle()
    runDocMonad userRes WindowsEnv <| 
        docMonad { 
            let! w1 = sourceWordFile "sample.docx" 
            let! relpath = getPathSuffix w1
            return relpath
        }


