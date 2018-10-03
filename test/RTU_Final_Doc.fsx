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
#I @"..\packages\Magick.NET-Q8-AnyCPU.7.8.0\lib\net40"
#r @"Magick.NET-Q8-AnyCPU.dll"
open ImageMagick

// Use FSharp.Data for CSV output (Proprietry.fs)
#I @"..\packages\FSharp.Data.3.0.0-beta3\lib\net45"
#r @"FSharp.Data.dll"

// Use ExcelProvider to read SAI numbers spreadsheet (Proprietry.fs)
#I @"..\packages\ExcelProvider.1.0.1\lib\net45"
#r "ExcelProvider.Runtime.dll"

#I @"..\packages\ExcelProvider.1.0.1\typeproviders\fsharp41\net45"
#r "ExcelDataReader.DataSet.dll"
#r "ExcelDataReader.dll"
#r "ExcelProvider.DesignTime.dll"
open FSharp.Interop.Excel


open System.IO

#load "..\src\DocMake\Base\Common.fs"
#load "..\src\DocMake\Base\FakeLike.fs"
#load "..\src\DocMake\Base\ImageMagickUtils.fs"
#load "..\src\DocMake\Base\OfficeUtils.fs"
#load "..\src\DocMake\Base\SimpleDocOutput.fs"
#load "..\src\DocMake\Builder\BuildMonad.fs"
#load "..\src\DocMake\Builder\Document.fs"
#load "..\src\DocMake\Builder\Basis.fs"
#load "..\src\DocMake\Builder\ShellHooks.fs"
open DocMake.Base.Common
open DocMake.Base.FakeLike
open DocMake.Builder.BuildMonad
open DocMake.Builder.Document
open DocMake.Builder.Basis

#load "..\src\DocMake\Tasks\IOActions.fs"
#load "..\src\DocMake\Tasks\DocFindReplace.fs"
#load "..\src\DocMake\Tasks\XlsFindReplace.fs"
#load "..\src\DocMake\Tasks\DocToPdf.fs"
#load "..\src\DocMake\Tasks\XlsToPdf.fs"
#load "..\src\DocMake\Tasks\PptToPdf.fs"
#load "..\src\DocMake\Tasks\MdToDoc.fs"
#load "..\src\DocMake\Tasks\PdfConcat.fs"
#load "..\src\DocMake\Tasks\PdfRotate.fs"
#load "..\src\DocMake\Tasks\DocPhotos.fs"
#load "..\src\DocMake\FullBuilder.fs"
open DocMake.FullBuilder
open DocMake.Tasks

#load "Proprietry.fs"
open Proprietry



let _templateRoot       = @"G:\work\Projects\rtu\final-docs\__Templates"
let _inputRoot          = @"G:\work\Projects\rtu\final-docs\input\Year4-Batch1"
let _outputRoot         = @"G:\work\Projects\rtu\final-docs\output\year4-batch1"



let makeCoverMatches (siteName:string) (saiLookups:SaiLookups) : option<SearchList> =  
    match getSaiNumber siteName saiLookups with
    | None -> None
    | Some sai -> 
        Some <| 
            [ "#SITENAME",      siteName   
            ; "#SAINUM" ,       sai
            ]



let cover (siteName:string) : FullBuild<PdfDoc> = 
    buildMonad { 
        let templatePath = _templateRoot </> @"FIRMWARE_UPGRADE RTU Cover Sheet.docx"
        let! template = getTemplateDoc templatePath
        let docOutName = sprintf "%s cover-sheet.docx" (underscoreName siteName)
        let lookups = getSaiLookups ()
        match makeCoverMatches siteName lookups with
        | Some matches -> 
            let! d1 = docFindReplace matches template >>= renameTo docOutName 
            let! d2 = docToPdf d1
            return d2
        | None -> throwError "cover - no sai number" |> ignore
    }

/// Survey / Disposed of / Install
/// "*Survey*.docx", "Survey" might not be the last word of the file name


let processSheet1 (globPattern:string) (inputPath:string) : FullBuild<PdfDoc> =
    match tryFindExactlyOneMatchingFile globPattern inputPath with
    | None -> throwError (sprintf "No sheet matching '%s'" globPattern)
    | Some doc -> getDocument doc >>= docToPdf




let surveySheet (inputPath:string) : FullBuild<PdfDoc> =
    let glob = "*Survey*.doc*"
    processSheet1 glob inputPath

    

let installSheet (inputPath:string) : FullBuild<PdfDoc> =
    let glob = "*Site Works*.doc*"
    processSheet1 glob inputPath


let makePhotosDoc (docTitle:string) (jpegSourceDirectory:string) (pdfName:string) (subFolder:string) : FullBuild<PdfDoc> = 
    let opts : DocPhotos.DocPhotosOptions = 
        { DocTitle = Some docTitle; ShowFileName = true; CopyToSubDirectory = subFolder } 
    docPhotos opts [jpegSourceDirectory] >>= docToPdf >>= renameTo pdfName


let surveyPhotos (siteName:string) : FullBuild<PdfDoc> = 
    let jpegsDir = _inputRoot </> safeName siteName </> @"Survey_Photos"
    let pdfName = sprintf "%s survey-photos.pdf" (safeName siteName)
    printfn "Survey Photos: %s" pdfName
    makePhotosDoc "Survey Photos" jpegsDir pdfName @"survey_photos"

let installPhotos (siteName:string) : FullBuild<PdfDoc> = 
    let jpegsDir = _inputRoot </> safeName siteName </> @"Install_Photos"
    let pdfName = sprintf "%s install-photos.pdf" (safeName siteName)
    printfn "Install Photos: %s" pdfName
    makePhotosDoc "Install Photos" jpegsDir pdfName @"install_photos"


let buildScript1 (inputPath:string) : FullBuild<PdfDoc> = 
    let siteName    = slashName <| FileInfo(inputPath).Name
    let uscoreName  = underscoreName siteName
    localSubDirectory uscoreName <| 
        buildMonad { 
            do! IOActions.clean () >>. IOActions.createOutputDirectory ()
            let finalName   = sprintf "%s S3953 RTU Replacement.pdf" uscoreName
            let! p1 = cover siteName
            let! p2 = surveySheet inputPath
            let! p3 = surveyPhotos siteName
            let! p4 = installSheet inputPath
            let! p5 = installPhotos siteName
            let pdfs : PdfDoc list = [p1;p2;p3;p4;p5]
            let! (final:PdfDoc) = makePdf finalName     <| pdfConcat pdfs
            return final            
        }

let makeUploadRow (name:SiteName) (sai:SAINumber) : UploadRow = 
    let docTitle = 
        sprintf "%s S3953 RTU Replacement" (name.Replace("/", " "))
    let docName = 
        sprintf "%s S3953 RTU Replacement.pdf" (underscoreName name)
    UploadTable.Row(assetName = name,
                    assetReference = sai,
                    projectName = "RTU Asset Replacement",
                    projectCode = "S3953",
                    title = docTitle,
                    category = "O & M Manuals",
                    referenceNumber = "S3953", 
                    revision = "1",
                    documentName = docName,
                    documentDate = standardDocumentDate (),
                    sheetVolume = "" )

let uploadReceipt (dirList:string list) : FullBuild<unit> = 
    let makeSiteName (path:string) = 
        slashName <| System.IO.DirectoryInfo(path).Name

    let uploadHelper = 
        { new IUploadHelper
          with member this.MakeUploadRow name sai = makeUploadRow name sai }

    buildMonad { 
        let siteNames = List.map makeSiteName dirList
        do! makeUploadForm  uploadHelper siteNames
    }

let buildScript () : FullBuild<unit> = 
    buildMonad { 
        let childFolders = 
            System.IO.Directory.GetDirectories(_inputRoot) |> Array.toList
        do! mapMz buildScript1 childFolders
        do! uploadReceipt childFolders
        return () 
    }
    

let main () : unit = 
    let env = 
        { WorkingDirectory = _outputRoot
          PrintQuality = PrintQuality.PqScreen
          PdfQuality = PdfPrintQuality.PdfScreen }

    let appConfig : FullBuildConfig = 
        { GhostscriptPath = @"C:\programs\gs\gs9.15\bin\gswin64c.exe"
          PdftkPath = @"C:\programs\PDFtk Server\bin\pdftk.exe" 
          PandocPath = @"pandoc" } 

    runFullBuild env appConfig (buildScript ()) 
