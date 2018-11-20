// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

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
#I @"..\packages\Magick.NET-Q8-AnyCPU.7.8.0\lib\net40"
#r @"Magick.NET-Q8-AnyCPU.dll"

// MarkdownDoc (my library)
#I @"..\packages\__MY_LIBS__\lib\net45"
#r @"MarkdownDoc.dll"

#load "..\src\DocBuild\Base\Common.fs"
#load "..\src\DocBuild\Internal\CommonUtils.fs"
#load "..\src\DocBuild\Internal\RunProcess.fs"
#load "..\src\DocBuild\Internal\PdftkRotate.fs"
#load "..\src\DocBuild\Internal\ImageMagickUtils.fs"
#load "..\src\DocBuild\Internal\ExcelUtils.fs"
#load "..\src\DocBuild\Internal\WordUtils.fs"
#load "..\src\DocBuild\Objects\Document.fs"
#load "..\src\DocBuild\Objects\PdfDoc.fs"
#load "..\src\DocBuild\Objects\ExcelDoc.fs"
#load "..\src\DocBuild\Objects\WordDoc.fs"
#load "..\src\DocBuild\Objects\PowerPointDoc.fs"
#load "..\src\DocBuild\Objects\MarkdownDoc.fs"
#load "..\src\DocBuild\Objects\JpegDoc.fs"
#load "..\src\DocBuild\Extras\PhotoBook.fs"
open DocBuild
open DocBuild.Base



//let cover (siteName:string) (saiNumber:string) : PdfDoc = 
//    let docName = sprintf "%s cover-sheet.docx" (safeName siteName)
//    let word = wordDoc 

