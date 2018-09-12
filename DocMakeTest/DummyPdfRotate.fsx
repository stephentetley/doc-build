// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.Word"
#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Excel\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.Excel"
#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.PowerPoint\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.PowerPoint"
#I @"C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c"
#r "office"

#I @"..\packages\Magick.NET-Q8-AnyCPU.7.4.6\lib\net40"
#r @"Magick.NET-Q8-AnyCPU.dll"


#load "..\DocMake\DocMake\Base\Common.fs"
#load "..\DocMake\DocMake\Base\FakeLike.fs"
#load "..\DocMake\DocMake\Base\ImageMagickUtils.fs"
#load "..\DocMake\DocMake\Base\OfficeUtils.fs"
#load "..\DocMake\DocMake\Base\SimpleDocOutput.fs"
#load "..\DocMake\DocMake\Builder\BuildMonad.fs"
#load "..\DocMake\DocMake\Builder\Document.fs"
#load "..\DocMake\DocMake\Builder\Basis.fs"
#load "..\DocMake\DocMake\Builder\ShellHooks.fs"
open DocMake.Base.Common
open DocMake.Base.FakeLike
open DocMake.Builder.BuildMonad
open DocMake.Builder.Document


#load "..\DocMake\DocMake\Tasks\DocFindReplace.fs"
#load "..\DocMake\DocMake\Tasks\XlsFindReplace.fs"
#load "..\DocMake\DocMake\Tasks\DocToPdf.fs"
#load "..\DocMake\DocMake\Tasks\XlsToPdf.fs"
#load "..\DocMake\DocMake\Tasks\PptToPdf.fs"
#load "..\DocMake\DocMake\Tasks\MdToDoc.fs"
#load "..\DocMake\DocMake\Tasks\PdfConcat.fs"
#load "..\DocMake\DocMake\Tasks\PdfRotate.fs"
#load "..\DocMake\DocMake\Tasks\DocPhotos.fs"
#load "..\DocMake\DocMake\FullBuilder.fs"
open DocMake.FullBuilder
open DocMake.Tasks



let test01 () = 
    let env = 
        { WorkingDirectory = @"D:\coding\fsharp\DocMake\data"
          PrintQuality = PrintQuality.PqScreen
          PdfQuality = PdfPrintQuality.PdfPrint }

    let appConfig : FullBuildConfig = 
        { GhostscriptPath = @"C:\programs\gs\gs9.15\bin\gswin64c.exe"
          PdftkPath = @"C:\programs\PDFtk Server\bin\pdftk.exe" 
          PandocPath = @"pandoc" } 


    let procM = 
        buildMonad { 
            let! (input:PdfDoc) = breturn <| makeDocument @"G:\work\working\In1.pdf"
            let! _ =  pdfRotateEmbed [ (PdfRotate.rotationSinglePage 1 PoEast) ] input
            return ()
        }

    runFullBuild env appConfig procM
