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
open Microsoft.Office.Interop

// ImageMagick for DocPhotos
#I @"..\packages\Magick.NET-Q8-AnyCPU.7.4.6\lib\net40"
#r @"Magick.NET-Q8-AnyCPU.dll"
open ImageMagick


// FAKE dependencies are getting onorous...
#I @"..\packages\FAKE.5.0.0-rc016.225\tools"
#r @"FakeLib.dll"
#I @"..\packages\Fake.Core.Globbing.5.0.0-beta021\lib\net46"
#r @"Fake.Core.Globbing.dll"
#I @"..\packages\Fake.IO.FileSystem.5.0.0-rc017.237\lib\net46"
#r @"Fake.IO.FileSystem.dll"
#I @"..\packages\Fake.Core.Trace.5.0.0-rc017.237\lib\net46"
#r @"Fake.Core.Trace.dll"
#I @"..\packages\Fake.Core.Process.5.0.0-rc017.237\lib\net46"
#r @"Fake.Core.Process.dll"



#load @"DocMake\Base\Common.fs"
#load @"DocMake\Base\FakeFake.fs"
#load @"DocMake\Base\FakeExtras.fs"
#load @"DocMake\Base\ImageMagickUtils.fs"
#load @"DocMake\Base\OfficeUtils.fs"
#load @"DocMake\Base\SimpleDocOutput.fs"
#load @"DocMake\Builder\BuildMonad.fs"
#load @"DocMake\Builder\Basis.fs"
#load @"DocMake\Builder\WordBuilder.fs"
#load @"DocMake\Builder\ExcelBuilder.fs"
#load @"DocMake\Builder\PowerPointBuilder.fs"
#load @"DocMake\Builder\GhostscriptBuilder.fs"
#load @"DocMake\Builder\PdftkBuilder.fs"
open DocMake.Base.Common
open DocMake.Base.FakeFake
open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis

#load @"DocMake\Lib\DocFindReplace.fs"
#load @"DocMake\Lib\DocPhotos.fs"
#load @"DocMake\Lib\DocToPdf.fs"
#load @"DocMake\Lib\XlsToPdf.fs"
#load @"DocMake\Lib\PptToPdf.fs"
#load @"DocMake\Lib\PdfConcat.fs"
#load @"DocMake\FullBuilder.fs"
open DocMake.Lib
open DocMake.FullBuilder

let _templateRoot       = @"G:\work\Projects\usar\final-docs\__Templates"


let makeCoverMatches (siteName:string) : SearchList =          
    [ "#SITENAME",          siteName   
    ; "#SAINUMBER" ,        "SAI000ZZZZZ"
    ]



let cover (siteName:string) : FullBuild<PdfDoc> = 
    buildMonad { 
        let templatePath = _templateRoot </> @"USAR Cover Sheet.docx"
        let! template = getTemplate templatePath
        let docOutName = sprintf "%s cover-sheet.docx" (safeName siteName)
        let matches = makeCoverMatches siteName
        let! d1 = docFindReplace matches template >>= renameTo docOutName 
        let! d2 = docToPdf d1
        return d2
    }