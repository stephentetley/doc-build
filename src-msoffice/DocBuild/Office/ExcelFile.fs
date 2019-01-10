// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Office


[<RequireQualifiedAccess>]
module ExcelFile = 

    open System.IO

    // Open at .Interop rather than .Excel then the Excel API has to be qualified
    open Microsoft.Office.Interop

    open DocBuild.Base
    open DocBuild.Office
    open DocBuild.Office.Internal
    open DocBuild.Office.OfficeMonad


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
                    (outputFile:string) : OfficeMonad<PdfFile> = 
        officeMonad { 
            let pdfQuality = excelExportQuality quality
            let! _ = execExcel <| fun app -> 
                        excelExportAsPdf app src.Path fitWidth pdfQuality outputFile
            let! pdf = liftDocMonad (getPdfFile outputFile)
            return pdf
        }

    let exportPdf (src:ExcelFile) 
                  (fitWidth:bool) 
                  (quality:PrintQuality) : OfficeMonad<PdfFile> = 
        let outputFile = Path.ChangeExtension(src.Path, "pdf")
        exportPdfAs src fitWidth quality outputFile


    // ************************************************************************
    // Find and replace

    let findReplaceAs (src:ExcelFile) (searches:SearchList) (outputFile:string) : OfficeMonad<ExcelFile> = 
        officeMonad { 
            let! ans = 
                execExcel <| fun app -> 
                        excelFindReplace app src.Path outputFile searches
            let! xlsx = liftDocMonad (getExcelFile outputFile)
            return xlsx
        }



    let findReplace (src:ExcelFile) (searches:SearchList)  : OfficeMonad<ExcelFile> = 
        findReplaceAs src searches src.NextTempName 

