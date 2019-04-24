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

    let execExcel (mf: Excel.Application -> DocMonad<#IExcelHandle,'a>) : DocMonad<#IExcelHandle,'a> = 
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
                    (outputAbsPath:string)
                    (src:ExcelDoc) : DocMonad<#IExcelHandle,PdfDoc> = 
        docMonad { 
            do! assertIsWorkingPath outputAbsPath
            let! pdfQuality = 
                asks (fun env -> env.PrintOrScreen) |>> excelExportQuality
            let! _ = 
                execExcel <| fun app -> 
                        liftResult (excelExportAsPdf app fitWidth pdfQuality src.AbsolutePath outputAbsPath)
            return! getWorkingPdfDoc outputAbsPath
        }

    /// Saves the file in the top-level working directory.
    let exportPdf (fitWidth:bool) 
                  (src:ExcelDoc) : DocMonad<#IExcelHandle,PdfDoc> = 
        docMonad { 
            let! path1 = extendWorkingPath src.FileName
            let outputAbsPath = Path.ChangeExtension(path1, "pdf")
            return! exportPdfAs fitWidth outputAbsPath src
        }



    // ************************************************************************
    // Find and replace

    let findReplaceAs (searches:SearchList) 
                      (outputAbsPath:string) 
                      (src:ExcelDoc) : DocMonad<#IExcelHandle,ExcelDoc> = 
        docMonad { 
            do! assertIsWorkingPath outputAbsPath
            let! ans = 
                execExcel <| fun app -> 
                        liftResult (excelFindReplace app searches src.AbsolutePath outputAbsPath)
            return! getWorkingExcelDoc outputAbsPath
        }



    let findReplace (searches:SearchList) 
                    (src:ExcelDoc) : DocMonad<#IExcelHandle,ExcelDoc> = 
        findReplaceAs searches src.AbsolutePath src
