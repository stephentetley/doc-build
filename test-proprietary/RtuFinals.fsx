// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


#r "netstandard"
open System
open System.Text.RegularExpressions

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
#I @"C:\Users\stephen\.nuget\packages\slformat\1.0.2-alpha-20190322\lib\netstandard2.0"
#r @"SLFormat.dll"
#I @"C:\Users\stephen\.nuget\packages\markdowndoc\1.0.1-alpha-20190508\lib\netstandard2.0"
#r @"MarkdownDoc.dll"


#load "..\src\DocBuild\Base\Internal\FilePaths.fs"
#load "..\src\DocBuild\Base\Internal\GhostscriptPrim.fs"
#load "..\src\DocBuild\Base\Internal\PandocPrim.fs"
#load "..\src\DocBuild\Base\Internal\PdftkPrim.fs"
#load "..\src\DocBuild\Base\Internal\ImageMagickPrim.fs"
#load "..\src\DocBuild\Base\Common.fs"
#load "..\src\DocBuild\Base\DocMonad.fs"
#load "..\src\DocBuild\Base\Document.fs"
#load "..\src\DocBuild\Base\Collection.fs"
#load "..\src\DocBuild\Base\FindFiles.fs"
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
open DocBuild.Office

#load "ExcelProviderHelper.fs"
open ExcelProviderHelper

// ImageMagick Dll loader.
// A hack to get over Dll loading error due to the 
// native dll `Magick.NET-Q8-x64.Native.dll`
[<Literal>] 
let NativeMagick = @"C:\Users\stephen\.nuget\packages\magick.net-q8-anycpu\7.9.2\runtimes\win-x64\native"
Environment.SetEnvironmentVariable("PATH", 
    Environment.GetEnvironmentVariable("PATH") + ";" + NativeMagick
    )



let WindowsEnv : DocBuildEnv = 
    { WorkingDirectory  = @"G:\work\Projects\rtu\final-docs\output\year4-rtu-batch2"
      SourceDirectory   = @"G:\work\Projects\rtu\final-docs\input\Year4-RTU-Batch02"
      IncludeDirectories = [ @"G:\work\Projects\rtu\final-docs\include" ]
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


type WorkTable = 
    ExcelFile< 
        FileName = @"G:\work\Projects\rtu\final-docs\input\Year4-RTU-Batch02\year4-docs-lookups.xlsx",
        ForceString = true >

type WorkRow = WorkTable.Row

type WorkItems = Map<string, WorkRow>

let readWorkSpeadsheet () : WorkItems = 
    let helper = 
        { new IExcelProviderHelper<WorkTable,WorkRow>
          with member this.ReadTableRows table = table.Data 
               member this.IsBlankRow row = match row.GetValue(0) with null -> true | _ -> false }
         
    excelReadRowsAsList helper (new WorkTable())
        |> List.map (fun row -> (safeName row.``Site Name``, row))
        |> Map.ofList

let renderMarkdownDoc (docTitle:string)
                      (markdown:MarkdownDoc) : DocMonadWord<PdfDoc> =
    docMonad {
        let! docx = Markdown.markdownToWord markdown
        return! WordDocument.exportPdf  docx |>> setTitle docTitle
    }


let coverSeaches (row:WorkRow) : SearchList = 
    [ ("#SAINUM", row.``Sai Number``)
    ; ("#SITENAME", row.``Site Name``) 
    ; ("#YEAR", row.``Year ``) ]


let genCover (workRow:WorkRow) : DocMonadWord<PdfDoc> = 
    let outputName = sprintf "%s cover.docx" (workRow.``Site Name`` |> safeName )
    let searches : SearchList = coverSeaches workRow
    docMonad { 
        let! (template:WordDoc) = getIncludeWordDoc "TEMPLATE MM3x-to-MMIM Cover Sheet.docx"
        let! outpath = extendWorkingPath outputName
        let! wordFile = WordDocument.findReplaceAs searches outpath template
        return! WordDocument.exportPdf wordFile
    }

let sourceWordDocToPdf (folder1:string) (fileGlob:string) (row:WorkRow) :DocMonadWord<PdfDoc option> = 
    let subdirectory =  (row.``Site Name`` |> safeName ) </> folder1
    localSourceSubdirectory (subdirectory) 
        <| docMonad { 
            printfn "<<<<<"
            let! input = tryExactlyOne <<| findSourceFilesMatching fileGlob false 
            printfn ">>>>>"
            match input with
            | None -> return None
            | Some infile ->
                let! doc = getWordDoc infile
                return! (WordDocument.exportPdf doc |>> Some)
        }


let genSurvey (row:WorkRow) :DocMonadWord<PdfDoc option> = 
    sourceWordDocToPdf "1.Survey" "*urvey*.doc*" row
    

let genSiteWorks (row:WorkRow) :DocMonadWord<PdfDoc> = 
    optionToFailM "No Site Works document" 
                  (sourceWordDocToPdf "2.Site_work" "*Works*.doc*" row)
                

let genSurveyPhotos (row:WorkRow) : DocMonadWord<PdfDoc option> = 
    let name1 = safeName row.``Site Name``
    let props : PandocWordShim.PhotoBookConfig = 
        { Title = "Survey Photos"
        ; SourceSubdirectory = name1 </> "1.Survey" </> "photos"
        ; WorkingSubdirectory = "survey_photos"
        ; RelativeOutputName = sprintf "%s survey photos.md" name1 }
    optionMaybeM (PandocWordShim.makePhotoBook props)


let genWorkPhotos (row:WorkRow) : DocMonadWord<PdfDoc option> = 
    let name1 = safeName row.``Site Name``
    let props : PandocWordShim.PhotoBookConfig = 
        { Title = "Install Photos"
        ; SourceSubdirectory  = name1 </> "2.Site_work" </> "photos"
        ; WorkingSubdirectory = "install_photos"
        ; RelativeOutputName= sprintf "%s install photos.md" name1 }
    optionMaybeM (PandocWordShim.makePhotoBook props)
    

let build1 (dict : WorkItems) : DocMonadWord<PdfDoc> = 
    docMonad { 
        let! name1 = sourceDirectoryName ()
        let  safeSiteName = name1 |> safeName
        let! row = liftOption "Could Not find row" (Map.tryFind name1 dict)
        let! cover = genCover row 
        let! oSurvey = genSurvey row
        let! works = genSiteWorks row
        let! oSurveyPhotos = genSurveyPhotos row
        let! oWorksPhotos = genWorkPhotos row

        let (col:PdfCollection) = 
            Collection.singleton cover 
                &^^ oSurvey     &^^ oSurveyPhotos
                &^^ works       &^^ oWorksPhotos

        let finalName = sprintf "%s Final.pdf" safeSiteName |> safeName
        return! Pdf.concatPdfs Pdf.GsDefault finalName col 
    }


let isLike (pattern:string) (source:string) = 
    Regex.IsMatch(input=source, pattern=pattern)


let main () = 
    let sites : WorkItems = readWorkSpeadsheet () 
    let resources = WindowsWordResources ()
    let options = { defaultSkeletonOptions with TestingSample = TakeDirectories 5 }
    runDocMonad resources WindowsEnv 
        <| foreachSourceDirectory options (build1 sites)