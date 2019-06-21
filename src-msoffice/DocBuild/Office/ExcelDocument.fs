// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Office


[<RequireQualifiedAccess>]
module ExcelDocument = 

    open System.IO

    // Open at .Interop rather than .Excel then the Excel API has to be qualified
    open Microsoft.Office.Interop

    open DocBuild.Base
    open DocBuild.Base.DocMonad
    open DocBuild.Office.Internal

    type ExcelHandle = 
        val mutable private ExcelApplication : Excel.Application 

        new () = 
            { ExcelApplication = null }

        /// Opens a handle as needed.
        member x.ExcelExe : Excel.Application = 
            match x.ExcelApplication with
            | null -> 
                let excel1 = initExcel ()
                x.ExcelApplication <- excel1
                excel1
            | app -> app

        interface IResourceFinalize with
            member x.RunFinalizer = 
                match x.ExcelApplication with
                | null -> () 
                | app -> finalizeExcel app

        interface IExcelHandle with
            member x.ExcelAppHandle = x

    and IExcelHandle =
        abstract ExcelAppHandle : ExcelHandle

    let execExcel (mf: Excel.Application -> DocMonad<'a, #IExcelHandle>) : DocMonad<'a, #IExcelHandle> = 
        docMonad { 
            let! userRes = getUserResources ()
            let excelHandle = userRes.ExcelAppHandle
            return! mf excelHandle.ExcelExe
        }

    // ************************************************************************
    // Export


    /// PqScreen maps to mininum
    /// PqPrint maps to standard
    let excelExportQuality (quality:PrintQuality) : Excel.XlFixedFormatQuality = 
        match quality with
        | Screen -> Excel.XlFixedFormatQuality.xlQualityMinimum
        | Print -> Excel.XlFixedFormatQuality.xlQualityStandard


    let exportPdfAs (fitWidth:bool)
                    (outputRelName:string)
                    (source:ExcelDoc) : DocMonad<PdfDoc, #IExcelHandle> = 
        docMonad { 
            let! outputAbsPath = extendWorkingPath outputRelName
            let! sourceName = getDocumentFileName source
            let! pdfQuality = 
                asks (fun env -> env.PrintOrScreen) |>> excelExportQuality
            let! _ = 
                execExcel <| fun app -> 
                        liftOperationResult "exportPdfAs" 
                            (fun _ -> excelExportAsPdf app fitWidth pdfQuality sourceName outputAbsPath)
            return! getPdfDoc outputAbsPath
        }

    /// Saves the file in the top-level working directory.
    let exportPdf (fitWidth:bool) 
                  (source : ExcelDoc) : DocMonad<PdfDoc, #IExcelHandle> = 
        docMonad { 
            let! sourceName = getDocumentFileName source
            let fileName = Path.ChangeExtension(sourceName, "pdf")
            return! exportPdfAs fitWidth fileName source
        }



    // ************************************************************************
    // Find and replace

    let findReplaceAs (searches:SearchList) 
                      (outputRelName:string) 
                      (source : ExcelDoc) : DocMonad<ExcelDoc, #IExcelHandle> = 
        docMonad { 
            let! outputAbsPath = extendWorkingPath outputRelName
            let! sourcePath = getDocumentPath source
            let! ans = 
                execExcel <| fun app -> 
                        liftOperationResult "findReplaceAs" 
                                (fun _ -> excelFindReplace app searches sourcePath outputAbsPath)
            return! getExcelDoc outputAbsPath
        }



    let findReplace (searches : SearchList) 
                    (source : ExcelDoc) : DocMonad<ExcelDoc, #IExcelHandle> = 
        docMonad { 
            let! sourceName = getDocumentFileName source
            return! findReplaceAs searches sourceName source
        }