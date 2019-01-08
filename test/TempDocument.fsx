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
#load "..\src\DocBuild\Base\Document.fs"
#load "..\src\DocBuild\Base\Collective.fs"
#load "..\src\DocBuild\Raw\Ghostscript.fs"
#load "..\src\DocBuild\Raw\Pandoc.fs"
#load "..\src\DocBuild\Raw\Pdftk.fs"
#load "..\src\DocBuild\Raw\ImageMagick.fs"
#load "..\src\DocBuild\Raw\PdftkRotate.fs"
#load "..\src\DocBuild\Document\Pdf.fs"
#load "..\src\DocBuild\Document\Jpeg.fs"
#load "..\src\DocBuild\Document\Markdown.fs"
#load "..\src\DocBuild\Extra\PhotoBook.fs"

#load "..\src-msoffice\DocBuild\Office\MsoExcel.fs"
#load "..\src-msoffice\DocBuild\Office\MsoWord.fs"
#load "..\src-msoffice\DocBuild\Office\WordDoc.fs"
#load "..\src-msoffice\DocBuild\Office\ExcelXls.fs"
#load "..\src-msoffice\DocBuild\Office\PowerPointPpt.fs"


open DocBuild
open DocBuild.Document.Markdown

open DocBuild.Raw.Ghostscript
open DocBuild.Base.Shell.Shell
open DocBuild.Document.Pdf

let getWorkingFile (name:string) = 
    let working = System.IO.Path.Combine(__SOURCE_DIRECTORY__, "..", "data")
    System.IO.Path.Combine(working, name)

let demo01 () = 
    let working = System.IO.Path.Combine(__SOURCE_DIRECTORY__, "..", "data")
    let gsOptions : ProcessOptions = 
        { WorkingDirectory = working
        ; ExecutableName = @"C:\programs\gs\gs9.15\bin\gswin64c.exe" 
        }

    let p1 = pdfFile <| getWorkingFile "One.pdf"
    let p2 = pdfFile <| getWorkingFile "Two.pdf"
    let cs1 = pdfColl [p1;p2]
    let outfile = getWorkingFile "Concat.pdf"

    ignore <| ghostscriptConcat gsOptions cs1 GsScreen outfile




