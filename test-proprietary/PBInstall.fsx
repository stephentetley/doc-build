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
#I @"C:\Users\stephen\.nuget\packages\markdowndoc\1.0.1-alpha-20190508\lib\netstandard2.0"
#r @"MarkdownDoc.dll"
open MarkdownDoc
open MarkdownDoc.Pandoc

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
#load "..\src\DocBuild\Base\Skeletons.fs"
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
open DocBuild.Document
open DocBuild.Document.Markdown
open DocBuild.Office


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
    { WorkingDirectory  = @"G:\work\Projects\events2\point-blue\output"
      SourceDirectory   = @"G:\work\Projects\events2\point-blue\batch4a_to_build"
      IncludeDirectories = [ @"G:\work\Projects\events2\point-blue\include" ]
      PrintOrScreen = PrintQuality.Screen
      PandocOpts = 
        { CustomStylesDocx = Some "custom-reference1.docx"
          PdfEngine = Some "pdflatex"
        }
      }

let WindowsWordResources () : AppResources<WordDocument.WordHandle> = 
    let userRes = new WordDocument.WordHandle()
    { GhostscriptExe = @"C:\programs\gs\gs9.15\bin\gswin64c.exe"
      PdftkExe = @"pdftk"
      PandocExe = @"pandoc"
      UserResources = userRes
    }

type DocMonadWord<'a> = DocMonad<'a, WordDocument.WordHandle>


let getSiteName (sourceName:string) : string = 
    sourceName.Replace("_", "/")


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
        | "T0975" -> 
            h1 (text "T0975 Event Duration Monitoring")
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
        let! (logo:JpegDoc) = getIncludeJpegDoc "YW-logo.jpg"
        let md = coverSheetMarkdown sai siteName phase logo.AbsolutePath
        let! outpath1 = extendWorkingPath "cover.md"
        let! mdDoc = saveMarkdown outpath1 md
        return! PandocWordShim.markdownToPdf mdDoc
    }

let genInstallSheet () : DocMonadWord<PdfDoc> = 
    docMonad { 
        do! askSourceDirectory () |>> fun o -> printfn "%s" (fileObjectName o)
        let! inputPath = optionToFailM "no install sheet" <| tryFindExactlyOneSourceFileMatching "*.docx" false
        let! wordDoc = getWordDoc inputPath
        return! WordDocument.exportPdfAs "install.pdf" wordDoc
        }
     
let build1 (phase:string) (saiMap:SaiMap) : DocMonadWord<PdfDoc> =        
    docMonad { 
        let! sourceName = askSourceDirectory () |>> fileObjectName
        let  siteName = getSiteName sourceName
        let! saiNumber = liftOption "No SAI Number" <| getSaiNumber saiMap siteName
        let! cover = genCoverSheet saiNumber siteName phase
        let! scope = genInstallSheet ()
        let col1 = Collection.fromList [ cover; scope ]  
        let finalName = sprintf "%s %s Final.pdf" sourceName phase |> safeName
        return! Pdf.concatPdfs Pdf.GsDefault finalName col1 
    }



let buildPhase (phase:string) (saiMap:SaiMap) : DocMonadWord<unit> =
    localSourceSubdirectory phase 
        << localWorkingSubdirectory phase
        <| foreachSourceIndividualOutput defaultSkeletonOptions (build1 phase saiMap)


let main () = 
    let resources = WindowsWordResources ()
    let saiMap : SaiMap = buildSaiMap () 
    runDocMonad resources WindowsEnv 
        <| docMonad { 
                // do! buildPhase "T0877" saiMap
                // do! buildPhase "T0942" saiMap
                do! buildPhase "T0975" saiMap
                return () 
            }


/// Have been observing a strange error where we see "zeroM" as the fail msg
/// rather than what is sent to "docError".
/// This doesn't provoke it...
let dummy01 () = 
    let resources = WindowsWordResources ()
    let saiMap : SaiMap = buildSaiMap () 
    runDocMonad resources WindowsEnv 
        <| docMonad { 
            let! saiCode    = liftOption "No SAI Number" <| getSaiNumber saiMap "UNKNOWN"
            let! goodNumber = liftOption "Error not expected" (Some 1)
            let! badNumber  = liftOption "Another Error" None
            return sprintf "%s-%i-%i" saiCode goodNumber badNumber
        }