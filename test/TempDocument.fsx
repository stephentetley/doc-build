// Copyright (c) Stephen Tetley 2018
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


#I @"C:\Users\stephen\.nuget\packages\markdowndoc\1.0.0\lib\netstandard2.0"
#r @"MarkdownDoc.dll"


#load "..\src\DocBuild\Base\Common.fs"
#load "..\src\DocBuild\Base\Shell.fs"
#load "..\src\DocBuild\Base\DocMonad.fs"
#load "..\src\DocBuild\Base\FakeLike.fs"
#load "..\src\DocBuild\Base\Document.fs"
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
#load "..\src-msoffice\DocBuild\Office\Common.fs"
#load "..\src-msoffice\DocBuild\Office\OfficeMonad.fs"
#load "..\src-msoffice\DocBuild\Office\WordFile.fs"
#load "..\src-msoffice\DocBuild\Office\ExcelFile.fs"
#load "..\src-msoffice\DocBuild\Office\PowerPointFile.fs"

open DocBuild.Base
open DocBuild.Document.Pdf
open DocBuild.Base.DocMonad

let getWorkingFile (name:string) = 
    let working = System.IO.Path.Combine(__SOURCE_DIRECTORY__, "..", "data")
    System.IO.Path.Combine(working, name)

let WindowsEnv : BuilderEnv = 
    { WorkingDirectory = System.IO.Path.Combine(__SOURCE_DIRECTORY__, "..", "data")
      GhostscriptExe = @"C:\programs\gs\gs9.15\bin\gswin64c.exe"
      PdftkExe = @"pdftk"
      PandocExe = @"pandoc"
      PandocReferenceDoc  = Some <| getWorkingFile "custom-reference1.docx"
    }


let demo01 () = 
    runDocMonad WindowsEnv <| 
        docMonad { 
            let! p1 = pdfFile <| getWorkingFile "One.pdf"
            let! p2 = pdfFile <| getWorkingFile "Two.pdf"
            let! p3 = pdfFile <| getWorkingFile "Three.pdf"
            let pdfs = [p1;p2;p3]
            let outfile = getWorkingFile "Concat.pdf"
            let! _ = ghostscriptConcat pdfs GsScreen outfile
            return ()
        }


let demo02 () = 
    runDocMonad WindowsEnv <| 
        docMonad { 
            let! p1 = pdfFile <| getWorkingFile "Concat.pdf"
            let! pageCount = pdfPageCount p1
            return pageCount
        }





