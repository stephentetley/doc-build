// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Document


[<AutoOpen>]
module ExcelXls = 


    // Open at .Interop rather than .Excel then the Word API has to be qualified
    open Microsoft.Office.Interop

    open DocBuild.Base
    open DocBuild.Base.Document
    open DocBuild.Raw.MsoExcel
    open DocBuild.Document.Pdf



    type ExcelExportQuality = 
        | ExcelQualityMinimum
        | ExcelQualityStandard



    let excelExportQuality (quality:ExcelExportQuality) : Excel.XlFixedFormatQuality = 
        match quality with
        | ExcelQualityMinimum -> Excel.XlFixedFormatQuality.xlQualityMinimum
        | ExcelQualityStandard -> Excel.XlFixedFormatQuality.xlQualityStandard


    type ExcelDoc = 
        val private ExcelDoc : Document

        new (filePath:string) = 
            { ExcelDoc = new Document(filePath = filePath) }



        member x.ExportAsPdf( fitWidth:bool
                            , quality:ExcelExportQuality
                            , outFile:string ) : unit = 
            // Don't make a temp file if we don't have to
            let srcFile = x.ExcelDoc.TempFile
            withExcelApp <| fun app -> 
                try 
                    let workbook : Excel.Workbook = app.Workbooks.Open(srcFile)
                    if fitWidth then 
                        workbook.Sheets 
                            |> Seq.cast<Excel.Worksheet>
                            |> Seq.iter (fun (sheet:Excel.Worksheet) -> 
                                sheet.PageSetup.Zoom <- false
                                sheet.PageSetup.FitToPagesWide <- 1)
                    else ()

                    workbook.ExportAsFixedFormat (Type=Excel.XlFixedFormatType.xlTypePDF,
                                                     Filename=outFile,
                                                     IncludeDocProperties=true,
                                                     Quality = excelExportQuality quality
                                                     )
                    workbook.Close (SaveChanges = false)
                with
                | ex -> failwith ex.Message


        member x.ExportAsPdf(fitWidth:bool, quality:ExcelExportQuality) : unit =
            // Don't make a temp file if we don't have to
            let srcFile = x.ExcelDoc.TempFile
            let outFile:string = System.IO.Path.ChangeExtension(srcFile, "pdf")
            x.ExportAsPdf(fitWidth = fitWidth, quality = quality, outFile = outFile)

        member x.SaveAs(outputPath: string) : unit = 
            x.ExcelDoc.SaveAs outputPath

        member x.FindReplace(searches:SearchList) : unit = 
            withExcelApp <| fun app -> 
                let tempFile = x.ExcelDoc.TempFile
                excelFindReplace app tempFile None searches


    let excelDoc (path:string) : ExcelDoc = new ExcelDoc (filePath = path)

