// Copyright (c) Stephen Tetley 2019
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
open MarkdownDoc
open MarkdownDoc.Pandoc

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
open DocBuild.Document.Markdown
open DocBuild.Office
open DocBuild.Office.PandocWordShim


#load "ExcelProviderHelper.fs"
#load "Proprietary.fs"
open Proprietary

// ImageMagick Dll loader.
// A hack to get over Dll loading error due to the 
// native dll `Magick.NET-Q8-x64.Native.dll`
[<Literal>] 
let NativeMagick = @"C:\Users\stephen\.nuget\packages\magick.net-q8-anycpu\7.9.2\runtimes\win-x64\native"
Environment.SetEnvironmentVariable("PATH", 
    Environment.GetEnvironmentVariable("PATH") + ";" + NativeMagick
    )

// Real code...



let WindowsEnv : DocBuildEnv = 
    { WorkingDirectory  = @"G:\work\Projects\events2\point-blue\batch3_to_build\output"
      SourceDirectory   = @"G:\work\Projects\events2\point-blue\batch3_to_build\input"
      IncludeDirectories = [ @"G:\work\Projects\events2\point-blue\batch3_to_build\include" ]
      PrintOrScreen = PrintQuality.Screen
      PandocOpts = 
        { CustomStylesDocx = Some "custom-reference1.docx"
          PdfEngine = Some "pdflatex"
        }
      }

let WindowsWordResources () : Resources<WordDocument.WordHandle> = 
    let userRes = new WordDocument.WordHandle()
    { GhostscriptExe = @"C:\programs\gs\gs9.15\bin\gswin64c.exe"
      PdftkExe = @"pdftk"
      PandocExe = @"pandoc"
      UserResources = userRes
    }

type DocMonadWord<'a> = DocMonad<WordDocument.WordHandle,'a>


let coverSheetMarkdown (sai:string) 
                       (siteName:string) 
                       (phase:string) 
                       (logoPath:string) : Markdown = 
    let title = 
        match phase with
        | "T0877" -> 
            h1 (text "T0877 Hawkeye 2 to Point Blue Asset Replacement (Phase 1)")
        | "T0942" -> 
            h1 (text "T0942 Hawkeye 2 to Point Blue Asset Replacement (Phase 2)")
        | _ -> 
            h1 (text "Error unknown Phase")
    concatMarkdown
        <|  [ markdownText (inlineImage "" logoPath None)
            ; nbsp ; nbsp
            ; title
            ; nbsp ; nbsp
            ; h2 (text sai ^+^ text siteName)
            ; nbsp ; nbsp
            ; h2 (text "Asset Replacement Project Partners")
            ; markdownText (doubleAsterisks (text "Metasphere") ^+^ text "Project Delivery")
            ; markdownText (doubleAsterisks (text "OnSite") ^+^ text "Installation and Commmissioning")
            ; nbsp ; nbsp
            ; h2 (text "Contents")
            ; markdown (unorderedList [paraText (text "Point Blue Installation / Commissioning Form")])
            ]

let genCoverSheet (sai:string) 
                  (siteName:string) 
                  (phase:string) : DocMonadWord<PdfDoc> = 
    docMonad {
        let! logo = includeJpegDoc "YW-logo.jpg"
        let md = coverSheetMarkdown sai siteName phase logo.AbsolutePath
        let! outpath1 = extendWorkingPath "cover.md"
        printfn "%O" md
        let! mdDoc = saveMarkdown outpath1 md
        return! markdownToWordToPdf mdDoc
    }

let genInstallSheet () : DocMonadWord<PdfDoc> = 
    docMonad { 
        do! askSourceDirectory () |>> fun o -> printfn "%s" (getPathName1 o)
        let! inputPath = optionFailM "no match" <| tryFindExactlyOneSourceFileMatching "*.docx" false
        let! wordDoc = getWordDoc inputPath
        let! outpath1 = extendWorkingPath "install.pdf"
        return! WordDocument.exportPdfAs outpath1 wordDoc
        }
     
let build1 (siteName:string) 
           (saiNumber:string)
           (phase:string) : DocMonadWord<PdfDoc> =
    let safe = safeName siteName           
    docMonad { 
        let! cover = genCoverSheet saiNumber siteName phase
        let! scope = genInstallSheet ()
        let col1 = Collection.fromList [ cover; scope ]  
        let! outputAbsPath = extendWorkingPath (sprintf "%s %s Final.pdf" safe phase)
        return! Pdf.concatPdfs Pdf.GsDefault col1 outputAbsPath 
    }


let getWorkList () : DocMonadWord<string list> = 
    askSourceDirectory () >>= fun srcDir -> 
    let dirs = System.IO.DirectoryInfo(srcDir).GetDirectories()
                    |> Array.map (fun info -> info.Name)
                    |> Array.toList
    mreturn dirs
 

let buildPhase (phase:string) (saiMap:SaiMap) : DocMonadWord<unit> =
    localSourceSubdirectory phase 
        <|  docMonad { 
                let! worklist = getWorkList () 
                do! forMz worklist <| fun dir -> 
                    let sitename = dir.Replace("_", "/")
                    let sai0 = getSaiNumber saiMap sitename 
                    let sai = 
                        match sai0 with 
                        | None -> printfn "BAD - %s" sitename; "SAI0000"
                        | Some ans -> ans
                    commonSubdirectory dir (build1 sitename sai phase)
                return ()
            }

let main () = 
    let resources = WindowsWordResources ()
    let saiMap : SaiMap = buildSaiMap () 
    runDocMonad resources WindowsEnv 
        <| docMonad { 
                do! buildPhase "T0877" saiMap
                do! buildPhase "T0942" saiMap
                return () 
            }
