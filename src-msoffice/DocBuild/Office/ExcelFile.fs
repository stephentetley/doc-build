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


    let exportPdfAs (quality:PrintQuality) 
                    (fitWidth:bool)
                    (outputName:string)
                    (src:ExcelFile) : DocMonad<#HasExcelHandle,PdfFile> = 
        docMonad { 
            let! outputPath = getOutputPath outputName
            let pdfQuality = excelExportQuality quality
            let! _ = 
                execExcel <| fun app -> 
                        liftResult (excelExportAsPdf app  fitWidth pdfQuality src.AbsolutePath outputPath)
            let! pdf = workingPdfFile outputName
            return pdf
        }

    /// Saves the file in the working directory.
    let exportPdf (quality:PrintQuality)
                  (fitWidth:bool) 
                  (src:ExcelFile) : DocMonad<#HasExcelHandle,PdfFile> = 
        docMonad { 
            let! local = Path.GetFileName(src.AbsolutePath) |> changeToWorkingFile
            let outputFile = Path.ChangeExtension(local.AbsolutePath, "pdf")
            let! pdf = exportPdfAs quality fitWidth outputFile src
            return pdf
        }
        /// exportPdfAs src fitWidth quality outputFile


    // ************************************************************************
    // Find and replace

    let findReplaceAs (searches:SearchList) (outputName:string) (src:ExcelFile) : DocMonad<#HasExcelHandle,ExcelFile> = 
        docMonad { 
            let! outputPath = getOutputPath outputName
            let! ans = 
                execExcel <| fun app -> 
                        liftResult (excelFindReplace app searches src.AbsolutePath outputPath)
            let! xlsx = workingExcelFile outputPath
            return xlsx
        }



    let findReplace (searches:SearchList) (src:ExcelFile) : DocMonad<#HasExcelHandle,ExcelFile> = 
        findReplaceAs searches src.FileName src
