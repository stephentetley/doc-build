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
            return! mf excelHandle.ExcelExe
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
                    (outputAbsPath:string)
                    (src:ExcelFile) : DocMonad<#HasExcelHandle,PdfFile> = 
        docMonad { 
            do! assertIsWorkingPath outputAbsPath
            let pdfQuality = excelExportQuality quality
            let! _ = 
                execExcel <| fun app -> 
                        liftResult (excelExportAsPdf app  fitWidth pdfQuality src.LocalPath outputAbsPath)
            return! workingPdfFile outputAbsPath
        }

    /// Saves the file in the top-level working directory.
    let exportPdf (quality:PrintQuality)
                  (fitWidth:bool) 
                  (src:ExcelFile) : DocMonad<#HasExcelHandle,PdfFile> = 
        docMonad { 
            let! path1 = extendWorkingPath src.FileName
            let outputAbsPath = Path.ChangeExtension(path1, "pdf")
            return! exportPdfAs quality fitWidth outputAbsPath src
        }



    // ************************************************************************
    // Find and replace

    let findReplaceAs (searches:SearchList) (outputAbsPath:string) (src:ExcelFile) : DocMonad<#HasExcelHandle,ExcelFile> = 
        docMonad { 
            do! assertIsWorkingPath outputAbsPath
            let! ans = 
                execExcel <| fun app -> 
                        liftResult (excelFindReplace app searches src.LocalPath outputAbsPath)
            return! workingExcelFile outputAbsPath
        }



    let findReplace (searches:SearchList) (src:ExcelFile) : DocMonad<#HasExcelHandle,ExcelFile> = 
        findReplaceAs searches src.LocalPath src
