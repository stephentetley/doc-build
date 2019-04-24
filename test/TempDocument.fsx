// Copyright (c) Stephen Tetley 2018,2019
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
#I @"C:\Users\stephen\.nuget\packages\Magick.NET-Q8-AnyCPU\7.11.1\lib\netstandard20"
#r @"Magick.NET-Q8-AnyCPU.dll"

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
open DocBuild.Document

open DocBuild.Office

let WindowsEnv : DocBuildEnv = 
    let cwd = System.IO.Path.Combine(__SOURCE_DIRECTORY__, "..", "data") 
    { WorkingDirectory = cwd
      SourceDirectory = cwd
      IncludeDirectories = [ cwd </> "include" ]
      PrintOrScreen = PrintQuality.Screen
      PandocOpts = 
        { CustomStylesDocx = None
          PdfEngine = Some "pdflatex"
        }
    }
let makeResources (userRes:'res) : AppResources<'res> = 
    { GhostscriptExe = @"C:\programs\gs\gs9.15\bin\gswin64c.exe"
      PdftkExe = @"pdftk"
      PandocExe = @"pandoc" 
      UserResources = userRes
    }


let demo01 () = 
    runDocMonadNoCleanup (makeResources ()) WindowsEnv <| 
        docMonad { 
            let! p1 = getWorkingPdfDoc "One.pdf"
            let! p2 = getWorkingPdfDoc "Two.pdf" 
            let! p3 = getWorkingPdfDoc "Three.pdf"
            let pdfs = Collection.fromList [p1;p2;p3]
            let! outfile = extendWorkingPath "Concat.pdf"
            let! _ = Pdf.concatPdfs Pdf.GsScreen pdfs outfile 
            return ()
        }


let demo02 () = 
    runDocMonadNoCleanup (makeResources ()) WindowsEnv <| 
        docMonad { 
            let! p1 = getWorkingPdfDoc "Concat.pdf"
            let! pageCount = Pdf.countPages p1
            return pageCount
        }



let demo03 () = 
    let resources = makeResources <| new WordDocument.WordHandle()
    runDocMonad resources WindowsEnv <| 
        docMonad { 
            let! w1 = getWorkingWordDoc "sample.docx" 
            return! WordDocument.exportPdf w1 
        }

let demo04 () = 
    let resources = makeResources <| new WordDocument.WordHandle()
    runDocMonad resources WindowsEnv <| 
        docMonad { 
            let! w1 = getSourceWordDoc "sample.docx" 
            return w1.AbsolutePath
        }


let demo05 () = 
    let resources = makeResources <| new WordDocument.WordHandle()
    runDocMonad resources WindowsEnv <| 
        docMonad { 
            return! findAllSourceFilesMatching "*.pdf" true
        }

let demo06 () = 
    let resources = makeResources <| new WordDocument.WordHandle()
    runDocMonad resources WindowsEnv <| 
        assertIsSourcePath @"D:\coding\fsharp\doc-build\data\Concat.pdf"

let demo06a () = 
    let resources = makeResources <| new WordDocument.WordHandle()
    runDocMonad resources WindowsEnv <| 
        (askSourceDirectory () |>> fun (src:string) -> Internal.FilePaths.pathToSegments src)



