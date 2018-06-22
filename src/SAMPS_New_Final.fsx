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


#I @"..\packages\ExcelProvider.0.8.2\lib"
#r "ExcelProvider.dll"
open FSharp.ExcelProvider

#I @"..\packages\Magick.NET-Q8-AnyCPU.7.4.6\lib\net40"
#r @"Magick.NET-Q8-AnyCPU.dll"
open ImageMagick


open System.IO
open System.Text.RegularExpressions



#load @"DocMake\Base\Common.fs"
#load @"DocMake\Base\FakeLike.fs"
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
open DocMake.Base.FakeLike
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



let _templateRoot       = @"G:\work\Projects\samps\final-docs\__Templates"
let _inputRoot          = @"G:\work\Projects\samps\final-docs\input\June2018_batch01"
let _outputRoot         = @"G:\work\Projects\samps\final-docs\output\June_Batch01"


type UidTable = 
    ExcelFile< @"G:\work\Projects\samps\uid_numbers.xlsx",
               SheetName = "Sheet1",
               ForceString = true >

type UidRow = UidTable.Row

let uidsTableDict : ExcelProviderHelperDict<UidTable, UidRow> = 
    { GetRows     = fun imports -> imports.Data 
      NotNullProc = fun row -> match row.GetValue(0) with | null -> false | _ -> true }

let uidDict : Map<string,string> = 
    excelTableGetRows uidsTableDict (new UidTable()) 
        |> List.map (fun (row:UidRow) -> row.``Common Name``, row.UID)
        |> Map.ofList


let clean () : FullBuild<unit> = deleteWorkingDirectory () 
let outputDirectory () : FullBuild<unit> = createWorkingDirectory ()


let makeCoverMatches (siteName:string) : SearchList = 
    let sai = 
        match uidDict.TryFind siteName with
        | Some uid -> uid
        | None -> failwith "makeCoverMatches"            
    [ "#SITENAME",          siteName   
    ; "#SAINUMBER" ,        sai
    ]


let cover (siteName:string) : FullBuild<PdfDoc> = 
    buildMonad { 
        let templatePath = _templateRoot </> @"TEMPLATE Samps Cover Sheet.docx"
        let! template = getTemplate templatePath
        let docOutName = sprintf "%s cover-sheet.docx" (safeName siteName)
        let matches = makeCoverMatches siteName
        let! d1 = docFindReplace matches template >>= renameTo docOutName 
        let! d2 = docToPdf d1
        return d2
    }

// One survey sheet per site (even if multiple samplers)
let surveySheet (siteName:string) : FullBuild<PdfDoc> = 
    let inputSubDir = _inputRoot </> safeName siteName </> @"SURVEY"
    let outName = sprintf "%s sampler-survey.pdf" (safeName siteName) 
    match tryFindExactlyOneMatchingFile "*Sampler survey.xls*" inputSubDir  with
    | None -> throwError "No survey sheet"
    | Some xls -> getDocument xls >>= xlsToPdf true >>= renameTo outName
    
// One survey sheet per site (even if multiple samplers)
let surveyPres (siteName:string) : FullBuild<PdfDoc> = 
    let inputSubDir = _inputRoot </> safeName siteName </> @"SURVEY"
    let outName = sprintf "%s survey-presentation.pdf" (safeName siteName) 
    match tryFindExactlyOneMatchingFile "*.ppt*" inputSubDir  with
    | None -> throwError "No survey presentation"
    | Some ppt -> 
        printfn "PPT: '%s" ppt
        getDocument ppt >>= pptToPdf >>= renameTo outName


let makePhotosDoc (docTitle:string) (jpegSourceDirectory:string) (pdfName:string) (subFolder:string) : FullBuild<PdfDoc> = 
    let opts : DocPhotos.DocPhotosOptions = 
        { DocTitle = Some docTitle; ShowFileName = true; CopyToSubDirectory = subFolder } 
    docPhotos opts [jpegSourceDirectory] >>= docToPdf >>= renameTo pdfName


let surveyPhotos (siteName:string) : FullBuild<PdfDoc> = 
    let jpegsDir = _inputRoot </> safeName siteName </> @"SURVEY" </> @"PHOTOS"
    let pdfName = sprintf "%s survey-photos.pdf" (safeName siteName)
    printfn "Survey Photos: %s" pdfName
    makePhotosDoc "Survey Photos" jpegsDir pdfName @"survey_photos"

// copy-pdf
let citCircuitDiagram (siteName:string) : FullBuild<PdfDoc> = 
    let inputSubDir = _inputRoot </> safeName siteName </> @"CIT"
    let outName = sprintf "%s circuit-diagram.pdf" (safeName siteName) 
    match tryFindExactlyOneMatchingFile "*Circuit Diagram.pdf" inputSubDir  with
    | None -> throwError "No survey sheet"
    | Some pdf -> copyToWorkingDirectory pdf >>= renameTo outName

// xls-to-pdf
let citWorkbook (siteName:string) : FullBuild<PdfDoc> =     
    let inputSubDir = _inputRoot </> safeName siteName </> @"CIT"
    let outName = sprintf "%s cit-workbook.pdf" (safeName siteName) 
    match tryFindExactlyOneMatchingFile "*YW Workbook.xls*" inputSubDir  with
    | None -> throwError "No survey sheet"
    | Some xls -> getDocument xls >>= xlsToPdf true >>= renameTo outName


// May be more-than-one.
// doc-to-pdf OR copy-pdf
let installSheets (siteName:string) : FullBuild<PdfDoc list> =
    let makeOutName (inputFileName:string) : string = 
        let groups = Regex.Match(inputFileName, @"([A-z_]*) Sampler Replacement").Groups
        if groups.Count > 0 then 
            sprintf "%s %s sampler-install.pdf" (safeName siteName) (safeName (groups.Item(1).Value))
        else 
            sprintf "%s UNKNOWN sampler-install.pdf" (safeName siteName) 

    let inputSubDir = _inputRoot </> safeName siteName </> @"SITE_WORKS"
    match tryFindSomeMatchingFiles "*Replacement Record.doc*" inputSubDir  with
    | None -> 
        match tryFindSomeMatchingFiles "*Replacement Record.pdf" inputSubDir  with
        | None -> throwError "No install sheets"
        | Some xs -> 
            forM xs <| fun pdf -> copyToWorkingDirectory pdf >>= renameTo (makeOutName pdf)
    | Some xs -> 
        forM xs <| fun docx -> getDocument docx >>= docToPdf >>= renameTo (makeOutName docx)

// May be more-than-one.
// copy-pdf
let bottleMachine (siteName:string) : FullBuild<PdfDoc list> =
    let makeOutName (inputFilePath:string) : string = 
        let name = System.IO.FileInfo(inputFilePath).Name
        sprintf "%s %s sampler-install.pdf" (safeName siteName) name
    let inputSubDir = _inputRoot </> safeName siteName </> @"SITE_WORKS"
    match tryFindSomeMatchingFiles "*Bottle_Machine.pdf" inputSubDir  with
    | None -> breturn []
    | Some xs -> 
        forM xs <| fun pdf -> copyToWorkingDirectory pdf >>= renameTo (makeOutName pdf)


let installPhotos (siteName:string) : FullBuild<PdfDoc list> = 
    let jpegsDir = _inputRoot </> safeName siteName </> @"SITE_WORKS" </> @"PHOTOS"
    // let source = { InputDirectory = jpegsDir; RenameProc = None} : DocPhotos.JpegInputSource
    let pdfName = sprintf "%s install-photos.pdf" (safeName siteName)
    printfn "Install Photos: %s" pdfName
    attempt (makePhotosDoc "Install Photos" jpegsDir pdfName @"install_photos" |>> (fun a -> [a]))
        <|> breturn []



let buildScript (siteName:string) : FullBuild<unit> = 
    let subFolder = safeName siteName
    let finalName = sprintf "%s S3820 Sampler Asset Replacement.pdf" (safeName siteName)
    localSubDirectory subFolder <| 
        buildMonad { 
            do! clean () >>. outputDirectory () 
            let! d1 = cover siteName
            let! d2 = surveySheet siteName
            let! d3 = surveyPres siteName
            let! d4 = surveyPhotos siteName
            let! d5 = citCircuitDiagram siteName
            let! d6 = citWorkbook siteName
            let! ds1 = installSheets siteName
            let! ds2 = installPhotos siteName
            let! ds3 = bottleMachine siteName
            let pdfs = [d1; d2; d3; d4; d5; d6] @ ds1 @ ds2 @ ds3
            let! (final:PdfDoc) = makePdf finalName     <| pdfConcat pdfs

            return ()                
        }

let main () : unit = 
    let gsExe = @"C:\programs\gs\gs9.15\bin\gswin64c.exe"
    let pdftkExe = @"C:\programs\PDFtk Server\bin\pdftk.exe"
    let hooks = fullBuilderHooks gsExe pdftkExe

    let env = 
        { WorkingDirectory = _outputRoot
          PrintQuality = DocMakePrintQuality.PqScreen
          PdfQuality = PdfPrintSetting.PdfScreen }
    
    let proc : FullBuild<unit> = 
        let folders = 
            System.IO.Directory.GetDirectories(_inputRoot) |> Array.toList
        foriMz folders <|
            fun ix path -> 
                if ix >= 0 then 
                    printfn "Processing %i of %i... '%s'" (ix+1) folders.Length path
                    let name = System.IO.DirectoryInfo(path).Name |> fun s -> s.Replace("_", "/")
                    buildScript name
                else
                    breturn ()
    consoleRun env hooks proc

