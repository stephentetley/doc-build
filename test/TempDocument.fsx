﻿// Copyright (c) Stephen Tetley 2018,2019
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
#I @"C:\Users\stephen\.nuget\packages\markdowndoc\1.0.1-alpha-20190205\lib\netstandard2.0"
#r @"MarkdownDoc.dll"


open System

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
#load "..\src-msoffice\DocBuild\Office\WordFile.fs"
#load "..\src-msoffice\DocBuild\Office\ExcelFile.fs"
#load "..\src-msoffice\DocBuild\Office\PowerPointFile.fs"

open DocBuild.Base
open DocBuild.Document.Pdf
open DocBuild.Base.DocMonad
open DocBuild.Base.DocMonadOperators

open DocBuild.Office

let WindowsEnv : BuilderEnv = 
    let cwd = System.IO.Path.Combine(__SOURCE_DIRECTORY__, "..", "data") |> DirectoryPath
    { WorkingDirectory = cwd
      SourceDirectory = cwd
      IncludeDirectory = DirectoryPath (cwd <//> "include")
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
            return! WordFile.exportPdf PqScreen w1 
        }

let demo04 () = 
    let userRes = new WordFile.WordHandle()
    runDocMonad userRes WindowsEnv <| 
        docMonad { 
            let! w1 = sourceWordFile "sample.docx" 
            return! getDocPathSuffix w1
        }


let demo05 () = 
    let userRes = new WordFile.WordHandle()
    runDocMonad userRes WindowsEnv <| 
        docMonad { 
            return! findAllSourceFilesMatching "*.pdf" true
        }

let demo06 () = 
    let userRes = new WordFile.WordHandle()
    runDocMonad userRes WindowsEnv <| 
        assertIsSourcePath @"D:\coding\fsharp\doc-build\data\Concat.pdf"

let demo06a () = 
    let userRes = new WordFile.WordHandle()
    runDocMonad userRes WindowsEnv <| 
        (askSourceDirectory () |>> fun (src:DirectoryPath) -> src.Segments)



