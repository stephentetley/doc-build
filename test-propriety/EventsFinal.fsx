// Copyright (c) Stephen Tetley 2019
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


#load "..\src\DocBuild\Base\Common.fs"
#load "..\src\DocBuild\Base\Shell.fs"
#load "..\src\DocBuild\Base\DocMonad.fs"
#load "..\src\DocBuild\Base\DocMonadOperators.fs"
#load "..\src\DocBuild\Base\Document.fs"
#load "..\src\DocBuild\Base\FakeLike.fs"
#load "..\src\DocBuild\Base\FileIO.fs"
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
open DocBuild.Base.DocMonad
open DocBuild.Document
open DocBuild.Office

#load "Coversheet.fs"
open Coversheet


let inputRoot   = @"G:\work\Projects\events2\final-docs\input\CSO_SPS"
let outputRoot  = @"G:\work\Projects\events2\final-docs\output\CSO_SPS"

type DocMonadWord<'a> = DocMonad<WordFile.WordHandle, 'a>

let WindowsEnv : BuilderEnv = 
    { WorkingDirectory = @"G:\work\Projects\events2\final-docs\output\CSO_SPS"
      GhostscriptExe = @"C:\programs\gs\gs9.15\bin\gswin64c.exe"
      PdftkExe = @"pdftk"
      PandocExe = @"pandoc"
      PandocReferenceDoc  = 
        Some @"G:\work\Projects\events2\final-docs\input\include\custom-reference1.docx"
    }


let getSiteName (folderName:string) : string = 
    folderName.Replace('_', '/')


let getSaiNumber (siteName:string) : DocMonad<'res,string> = 
    dreturn "SAI00001234"       // TEMP


let coversheet (siteName:string) (saiNumber:string) : DocMonadWord<PdfFile> = 
    docMonad { 
        let logoPath = @"..\..\..\input" </> "include" </> "YW-logo.jpg"
        let! markdownFile = coversheet saiNumber siteName logoPath "coversheet.md"
        let! docx = Markdown.markdownToWord markdownFile
        let! pdf = WordFile.exportPdf docx PqScreen
        return pdf
    }



let getWorkList () : string list = 
    System.IO.DirectoryInfo(inputRoot).GetDirectories()
        |> Array.map (fun di -> di.Name)
        |> Array.toList

let buildOne (sourceName:string) : DocMonadWord<unit> = 
    localSubDirectory sourceName <| 
        docMonad {
            let siteName = getSiteName sourceName
            let! saiNumber = getSaiNumber siteName
            let! cover =  coversheet siteName saiNumber
            return ()
        }

let buildAll () : DocMonadWord<unit> = 
    let worklist = getWorkList ()
    foriMz worklist 
        <| fun ix name ->
                ignore <| printfn "Site %i: %s" (ix+1) name
                buildOne name

        


let demo01 () = 
    let userRes = new WordFile.WordHandle()
    runDocMonad userRes WindowsEnv 
        <| buildAll ()
