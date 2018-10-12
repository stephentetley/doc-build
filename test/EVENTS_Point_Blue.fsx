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


// ImageMagick for DocPhotos
#I @"..\packages\Magick.NET-Q8-AnyCPU.7.8.0\lib\net40"
#r @"Magick.NET-Q8-AnyCPU.dll"



#I @"..\packages\FSharp.Data.3.0.0-beta3\lib\net45"
#r @"FSharp.Data.dll"
open FSharp.Data


#I @"..\packages\ExcelProvider.1.0.1\lib\net45"
#r "ExcelProvider.Runtime.dll"

#I @"..\packages\ExcelProvider.1.0.1\typeproviders\fsharp41\net45"
#r "ExcelDataReader.DataSet.dll"
#r "ExcelDataReader.dll"
#r "ExcelProvider.DesignTime.dll"
open FSharp.Interop.Excel

#load "..\src\DocMake\Base\Common.fs"
#load "..\src\DocMake\Base\FakeLike.fs"
#load "..\src\DocMake\Base\ImageMagickUtils.fs"
#load "..\src\DocMake\Base\OfficeUtils.fs"
#load "..\src\DocMake\Base\SimpleDocOutput.fs"
#load "..\src\DocMake\Builder\BuildMonad.fs"
#load "..\src\DocMake\Builder\Document.fs"
#load "..\src\DocMake\Builder\Basis.fs"
#load "..\src\DocMake\Builder\ShellHooks.fs"
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

#load "Proprietary.fs"
open Proprietary


let _outputRoot     = @"G:\work\Projects\events2\point-blue\output"


// ****************************************************************************
// Generate output from a work list

type InputTable = 
    CsvProvider< @"G:\work\Projects\events2\point-blue\missing.csv",
                 HasHeaders = true >

type InputRow = InputTable.Row


let rows () : InputRow list  = (new InputTable()).Rows |> Seq.toList


let cover (newSiteName:string) (scheme:string) : FullBuild<PdfDoc> = 
    let coversLocation = 
        match scheme with
        | "T0877" -> @"G:\work\Projects\events2\point-blue\cover-sheets\phase1"
        | "T0942" -> @"G:\work\Projects\events2\point-blue\cover-sheets\phase2"
        // Look somewhere we won't find it...
        | _ -> @"G:\work\Projects\events2\point-blue\cover-sheets"

    let pattern1 = sprintf "%s*" (underscoreName newSiteName)
    printfn "Pattern: '%s'" pattern1
    buildMonad { 
        match tryFindExactlyOneMatchingFile pattern1 coversLocation with
        | Some doc -> 
            printfn "doc is '%s'" doc
            let! d1 = docToPdf (makeDocument doc)
            return d1
        | None -> 
            throwError "cover - no sheet" |> ignore
    }

let workSheet (oldSiteName:string) : FullBuild<PdfDoc> = 
    let formsLoc =  @"G:\work\Projects\events2\point-blue\pb-commissioning-forms"
    let pattern1 = sprintf "%s*" (underscoreName oldSiteName)
    buildMonad { 
        match tryFindExactlyOneMatchingFile pattern1 formsLoc with
        | Some doc -> 
            let! d1 = docToPdf (makeDocument doc)
            return d1
        | None -> 
            throwError "work-sheet - no sheet" |> ignore
    }

let makeFinal (newSiteName:string) (scheme:string) (cover:PdfDoc) (worksheet:PdfDoc) : FullBuild<PdfDoc> = 
    let finalName = sprintf "%s %s Hawkeye Asset Replacement.pdf" (safeName newSiteName) scheme
    buildMonad { 
        let pdfs : PdfDoc list = [cover; worksheet]
        printfn "Final: '%s'" finalName
        let! (final:PdfDoc) = makePdf finalName <| pdfConcat pdfs
        return final   
    }


let build1 (row:InputRow) : FullBuild<PdfDoc> = 
    let phaseFolder = 
        match row.Scheme with
        | "T0877" -> "phase1"
        | "T0942" -> "phase2"
        | _ -> "phase2"
    localSubDirectory phaseFolder
        <| buildMonad {
            let! d1 = cover row.``New AI2 Name`` row.Scheme
            let! d2 = workSheet row.Missing
            let! final = makeFinal row.``New AI2 Name`` row.Scheme d1 d2
            return final
        }

 

let makeUploadRow (name:SiteName) (sai:SAINumber) (projCode:string) : UploadRow = 
    let docTitle = 
        sprintf "%s %s Hawkeye Asset Replacement" (name.Replace("/", " ")) projCode
    let docName = 
        sprintf "%s %s Hawkeye Asset Replacement.pdf" (underscoreName name) projCode

    UploadTable.Row(assetName = name,
                    assetReference = sai,
                    projectName = "RTU Asset Replacement",
                    projectCode = projCode,
                    title = docTitle,
                    category = "O & M Manuals",
                    referenceNumber = projCode, 
                    revision = "1",
                    documentName = docName,
                    documentDate = standardDocumentDate (),
                    sheetVolume = "" )


let uploadReceipts (rows:InputRow list): FullBuild<unit> = 
    let filterScheme (scheme:string) : InputRow list = 
        List.filter (fun (row:InputRow) -> row.Scheme = scheme) rows

    let uploadHelper (projCode:string) = 
        { new IUploadHelper<InputRow>
          with member __.MakeUploadRow name sai = makeUploadRow name sai projCode
               member __.ToSiteRecord row = { SiteName = row.``New AI2 Name``
                                            ; Uid = row.``SAI number`` } }
    
    buildMonad { 
        do! localSubDirectory "phase1" (makeUploadForm  (uploadHelper "T0877") (filterScheme "T0877"))
        do! localSubDirectory "phase2" (makeUploadForm  (uploadHelper "T0942") (filterScheme "T0942"))
        return ()
    }       


let buildScript () : FullBuild<unit> = 
    buildMonad { 
        let siteList = rows () 
        let count = siteList.Length
        do! foriMz siteList <| fun ix row -> 
                printfn "Site %i of %i (%s):" (ix+1) count row.``New AI2 Name``
                fmapM ignore <| build1 row

        do! uploadReceipts siteList
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