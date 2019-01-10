// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Office


[<RequireQualifiedAccess>]
module ExcelFile = 

    open System.IO

    // Open at .Interop rather than .Excel then the Excel API has to be qualified
    open Microsoft.Office.Interop

    open DocBuild.Base.Common
    open DocBuild.Base.Document
    open DocBuild.Office
    open DocBuild.Office.Internal
    open DocBuild.Office.OfficeMonad


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
            let! pdf = liftDocMonad (pdfFile outputFile)
            return pdf
        }

    let exportPdf (src:ExcelFile) 
                  (fitWidth:bool) 
                  (quality:PrintQuality) : OfficeMonad<PdfFile> = 
        let outputFile = Path.ChangeExtension(src.Path, "pdf")
        exportPdfAs src fitWidth quality outputFile

        //member x.ExportAsPdf( fitWidth:bool
        //                    , quality:ExcelExportQuality
        //                    , outFile:string ) : unit = 
        //    // Don't make a temp file if we don't have to
        //    let srcFile = x.ExcelDoc.ActiveFile
        //    withExcelApp <| fun app -> 
        //        try 
        //            let workbook : Excel.Workbook = app.Workbooks.Open(srcFile)
        //            if fitWidth then 
        //                workbook.Sheets 
        //                    |> Seq.cast<Excel.Worksheet>
        //                    |> Seq.iter (fun (sheet:Excel.Worksheet) -> 
        //                        sheet.PageSetup.Zoom <- false
        //                        sheet.PageSetup.FitToPagesWide <- 1)
        //            else ()

        //            workbook.ExportAsFixedFormat (Type=Excel.XlFixedFormatType.xlTypePDF,
        //                                             Filename=outFile,
        //                                             IncludeDocProperties=true,
        //                                             Quality = excelExportQuality quality
        //                                             )
        //            workbook.Close (SaveChanges = false)
        //        with
        //        | ex -> failwith ex.Message


        //member x.ExportAsPdf(fitWidth:bool, quality:ExcelExportQuality) : unit =
        //    // Don't make a temp file if we don't have to
        //    let srcFile = x.ExcelDoc.ActiveFile
        //    let outFile:string = System.IO.Path.ChangeExtension(srcFile, "pdf")
        //    x.ExportAsPdf(fitWidth = fitWidth, quality = quality, outFile = outFile)






    let findReplaceAs (src:ExcelFile) (searches:SearchList) (outputFile:string) : OfficeMonad<ExcelFile> = 
        officeMonad { 
            let! ans = 
                execExcel <| fun app -> 
                        excelFindReplace app src.Path outputFile searches
            let! xlsx = liftDocMonad (excelFile outputFile)
            return xlsx
        }



    let findReplace (src:ExcelFile) (searches:SearchList)  : OfficeMonad<ExcelFile> = 
        findReplaceAs src searches src.NextTempName 

