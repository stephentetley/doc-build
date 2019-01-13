// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Office


[<RequireQualifiedAccess>]
module ExcelFile = 

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

        interface ResourceFinalize with
            member x.RunFinalizer = 
                match x.ExcelApplication with
                | null -> () 
                | app -> finalizeExcel app

        interface HasExcelHandle with
            member x.ExcelAppHandle = x

    and HasExcelHandle =
        abstract ExcelAppHandle : ExcelHandle

    let execExcel (mf: Excel.Application -> DocMonad<#HasExcelHandle,'a>) : DocMonad<#HasExcelHandle,'a> = 
        docMonad { 
            let! userRes = askUserResources ()
            let excelHandle = userRes.ExcelAppHandle
            let! ans = mf excelHandle.ExcelExe
            return ans
        }

    // ************************************************************************
    // Export


    /// PqScreen maps to mininum
    /// PqPrint maps to standard
    let excelExportQuality (quality:PrintQuality) : Excel.XlFixedFormatQuality = 
        match quality with
        | PqScreen -> Excel.XlFixedFormatQuality.xlQualityMinimum
        | PqPrint -> Excel.XlFixedFormatQuality.xlQualityStandard


    let exportPdfAs (src:ExcelFile) 
                    (fitWidth:bool)
                    (quality:PrintQuality) 
                    (outputFile:string) : DocMonad<#HasExcelHandle,PdfFile> = 
        docMonad { 
            let pdfQuality = excelExportQuality quality
            let! _ = 
                execExcel <| fun app -> 
                        liftResult (excelExportAsPdf app src.Path fitWidth pdfQuality outputFile)
            let! pdf = getPdfFile outputFile
            return pdf
        }

    let exportPdf (src:ExcelFile) 
                  (fitWidth:bool) 
                  (quality:PrintQuality) : DocMonad<#HasExcelHandle,PdfFile> = 
        let outputFile = Path.ChangeExtension(src.Path, "pdf")
        exportPdfAs src fitWidth quality outputFile


    // ************************************************************************
    // Find and replace

    let findReplaceAs (src:ExcelFile) (searches:SearchList) (outputFile:string) : DocMonad<#HasExcelHandle,ExcelFile> = 
        docMonad { 
            let! ans = 
                execExcel <| fun app -> 
                        liftResult (excelFindReplace app src.Path outputFile searches)
            let! xlsx = getExcelFile outputFile
            return xlsx
        }



    let findReplace (src:ExcelFile) (searches:SearchList)  : DocMonad<#HasExcelHandle,ExcelFile> = 
        findReplaceAs src searches src.NextTempName 

