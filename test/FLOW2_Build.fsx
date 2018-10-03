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
open DocMake.Base.Common
open DocMake.Base.FakeLike
open DocMake.Builder.BuildMonad
open DocMake.Builder.Document
open DocMake.Builder.Basis
open DocMake.FullBuilder
open DocMake.Tasks

#load "Proprietry.fs"
open Proprietry


let _inputRoot      = @"G:\work\Projects\flow2\final-docs\Input\Batch02"
let _outputRoot     = @"G:\work\Projects\flow2\final-docs\Output\Batch02"
let _templateRoot   = @"G:\work\Projects\flow2\final-docs\__Templates"


let makeCoverMatches (siteName:string) (saiLookups:SaiLookups) : option<SearchList> =  
    match getSaiNumber siteName saiLookups with
    | None -> None
    | Some sai -> 
        Some <| 
            [ "#SITENAME",      siteName   
            ; "#SAINUM" ,       sai
            ]




// This should be a mandatory task
let cover (siteName:string) : FullBuild<PdfDoc> = 
    buildMonad { 
        let templatePath = _templateRoot </> "TEMPLATE Flow Confirmation Phase2 Cover.docx"
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


// Common procedure for both survey and Install photos
let photosDoc (docTitle:string) (jpegSrcDirectory:string) (pdfName:string) : FullBuild<PdfDoc> = 
    let folderName = DirectoryInfo(jpegSrcDirectory).Name
    let photoOpts:DocPhotos.DocPhotosOptions = 
        { DocTitle = Some docTitle; ShowFileName = true; CopyToSubDirectory = folderName } 

    buildMonad { 
        let! d1 = docPhotos photoOpts [jpegSrcDirectory]
        let! d2 = docToPdf d1 >>= renameTo pdfName
        return d2
        } <|> breturn zeroDocument
    
/// Scope of works should be a Word file or a Pdf 
let scopeOfWorks (inputPath:string) : FullBuild<PdfDoc> = 
    let sowDoc = 
        match tryFindExactlyOneMatchingFile "*Scope ?f Works*.doc*" inputPath with
        | None -> throwError "NO SCOPE OF WORKS"
        | Some src -> 
            printfn "Found: %s" src
            getDocument src >>= docToPdf >>= renameTo "scope-of-works.pdf"

    let sowPdf = 
        match tryFindExactlyOneMatchingFile "*Scope ?f Works*.pdf" inputPath with
        | None -> throwError "NO SCOPE OF WORKS"
        | Some src -> 
            printfn "Found: %s" src
            copyToWorkingDirectory src >>= renameTo "scope-of-works.pdf"

    sowDoc <|> sowPdf <&?> "scope-of-works"


    



/// As-builts are in a subfolder ```As_builts```
/// As-builts names may cause problems for pdftk...
let asBuilts (inputPath) : FullBuild<PdfDoc list> = 
    let renamer (ix:int) = sprintf "cit-cad-drawing-%03i.pdf" ix

    let proc1 (ix:int) (srcPath:string) : FullBuild<PdfDoc> = 
        if System.Text.RegularExpressions.Regex.IsMatch(srcPath, "\.ccw\.pdf") then 
            copyToWorkingDirectory srcPath >>= pdfRotateAllCcw >>= renameTo (renamer (ix+1)) 
        else if System.Text.RegularExpressions.Regex.IsMatch(srcPath, "\.cw\.pdf") then 
            copyToWorkingDirectory srcPath >>= pdfRotateAllCw >>= renameTo (renamer (ix+1)) 
        else
            copyToWorkingDirectory srcPath >>= renameTo (renamer (ix+1)) 

    findAllMatchingFiles "*.pdf" (inputPath </> "As_builts")
        |> mapiM proc1
    


    // If this isn't thunkified it will launch Excel when the code is loaded in FSI
let installSheets (inputPath:string) : FullBuild<PdfDoc list> = 
    let pdfGen (glob:string) (warnMsg:string) (namer:Printf.StringFormat<int -> string>) : FullBuild<PdfDoc list> = 
        match findAllMatchingFiles glob inputPath with
        | [] -> 
            printfn "%s" warnMsg; breturn []
             
        | xs -> 
            foriM xs (fun ix path -> 
                let name1 = sprintf namer ix
                printfn "installSheet: %s" path
                getDocument path >>= xlsToPdf true >>= renameTo name1)
    
    buildMonad { 
        let! ds1 = pdfGen "*Flow meter*.xls*" "NO FLOW METER INSTALL SHEETS" "Flow-meter-%i03-install.pdf"
        let! ds2 = pdfGen "*Pressure inst*.xls*" "NO PRESSURE SENSOR INSTALL SHEET" "Pressure-sensor-%i03-install.pdf"
        return ds1 @ ds2
    }


// *******************************************************


let makeFinalPdfName (siteName:string) : string = 
    sprintf "%s S4102 Flow Confirmation Manual.pdf" (underscoreName siteName)

let makeUploadRow (name:SiteName) (sai:SAINumber) : UploadRow = 
    let docTitle = 
        sprintf "%s S4102 Flow Confirmation Manual" (name.Replace("/", " "))
    let docName = 
        sprintf "%s S4102 Flow Confirmation Manual.pdf" (underscoreName name)
    UploadTable.Row(assetName = name,
                    assetReference = sai,
                    projectName = "Flow Confirmation",
                    projectCode = "S4102",
                    title = docTitle,
                    category = "O & M Manuals",
                    referenceNumber = "S4102", 
                    revision = "1",
                    documentName = docName,
                    documentDate = standardDocumentDate (),
                    sheetVolume = "" )

let uploadReceipt (dirList:string list) : FullBuild<unit> = 
    let siteFromPath (path:string) = 
        slashName <| System.IO.DirectoryInfo(path).Name
        
    let uploadHelper = 
        { new IUploadHelper
          with member this.MakeUploadRow name sai = makeUploadRow name sai }

    buildMonad { 
        let siteNames = List.map siteFromPath dirList
        do! makeUploadForm uploadHelper siteNames
    }


/// inputPath points to site input directory 
let buildScript1 (inputPath:string) : FullBuild<PdfDoc> = 
    let siteName    = slashName <| DirectoryInfo(inputPath).Name
    let cleanName   = safeName siteName
    let finalPdf = makeFinalPdfName siteName
    localSubDirectory cleanName <|
        buildMonad { 
            do! IOActions.clean () >>. IOActions.createOutputDirectory ()
            let siteInputDir = _inputRoot </> underscoreName siteName
            let! p1 = cover siteName
            let surveyJpegsPath = siteInputDir </> "Survey_Photos"
            let! p2 = photosDoc "Survey Photos" surveyJpegsPath "survey-photos.pdf"
            let! p3 = scopeOfWorks siteInputDir 
            let! ps1 = asBuilts siteInputDir
            let! ps2 = installSheets siteInputDir
            let surveyJpegsPath = siteInputDir </> "Install_Photos"
            let! pZ = photosDoc "Install Photos" surveyJpegsPath "install-photos.pdf"
            let (pdfs:PdfDoc list) = p1 :: p2 :: p3 :: (ps1 @ ps2 @ [pZ])
            let! (final:PdfDoc) = pdfConcat pdfs >>= renameTo finalPdf
            return final                 
        }

/// Also make upload receipt...
let buildScript () : FullBuild<unit>  = 
    buildMonad { 
        let inputs = 
            System.IO.Directory.GetDirectories(_inputRoot) 
                |> Array.toList
        do! mapMz buildScript1 inputs
        do! uploadReceipt inputs
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

    runFullBuild env appConfig <| buildScript ()


