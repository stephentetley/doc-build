// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


module DocBuild.ExcelDoc


// Open at .Interop rather than .Excel then the Word API has to be qualified
open Microsoft.Office.Interop

open DocBuild.Internal.CommonUtils
open DocBuild.Internal.ExcelUtils
open DocBuild.Common
open DocBuild.PdfDoc



type ExcelExportQuality = 
    | ExcelQualityMinimum
    | ExcelQualityStandard



let excelExportQuality (quality:ExcelExportQuality) : Excel.XlFixedFormatQuality = 
    match quality with
    | ExcelQualityMinimum -> Excel.XlFixedFormatQuality.xlQualityMinimum
    | ExcelQualityStandard -> Excel.XlFixedFormatQuality.xlQualityStandard


type ExcelDoc = 
    val private SourcePath : string
    val private TempPath : string

    new (filePath:string) = 
        { SourcePath = filePath
        ; TempPath = getTempFileName filePath }

    member internal v.TempFile
        with get() : string = 
            if System.IO.File.Exists(v.TempPath) then
                v.TempPath
            else
                System.IO.File.Copy(v.SourcePath, v.TempPath)
                v.TempPath
    
    member internal v.Updated 
        with get() : bool = System.IO.File.Exists(v.TempPath)

    member v.ExportAsPdf(fitWidth:bool, quality:ExcelExportQuality, outFile:string) : PdfDoc = 
        // Don't make a temp file if we don't have to
        let srcFile = if v.Updated then v.TempPath else v.SourcePath
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
                pdfDoc outFile
            with
            | ex -> failwith ex.Message


    member v.ExportAsPdf(fitWidth:bool, quality:ExcelExportQuality) : PdfDoc =
        // Don't make a temp file if we don't have to
        let srcFile = if v.Updated then v.TempPath else v.SourcePath
        let outFile:string = System.IO.Path.ChangeExtension(srcFile, "pdf")
        v.ExportAsPdf(fitWidth = fitWidth, quality = quality, outFile = outFile)

    member v.FindReplace(searches:SearchList) : ExcelDoc = 
        withExcelApp <| fun app -> 
            let tempFile = v.TempFile
            excelFindReplace app tempFile None searches
        v

let excelDoc (path:string) : ExcelDoc = new ExcelDoc (filePath = path)