// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


module DocBuild.ExcelDoc


// Open at .Interop rather than .Excel then the Word API has to be qualified
open Microsoft.Office.Interop

open DocBuild.Internal.Common
open DocBuild.PdfDoc


let private withExcelApp (operation:Excel.Application -> 'a) : 'a = 
    let app = new Excel.ApplicationClass(Visible = true) :> Excel.Application
    app.DisplayAlerts <- false
    app.EnableEvents <- false
    let result = operation app
    app.DisplayAlerts <- true
    app.EnableEvents <- true
    app.Quit ()
    result

type ExcelExportQuality = 
    | ExcelQualityMinimum
    | ExcelQualityStandard



let excelExportQuality (quality:ExcelExportQuality) : Excel.XlFixedFormatQuality = 
    match quality with
    | ExcelQualityMinimum -> Excel.XlFixedFormatQuality.xlQualityMinimum
    | ExcelQualityStandard -> Excel.XlFixedFormatQuality.xlQualityStandard


type ExcelDoc = 
    val XlsPath : string

    new (filePath:string) = 
        { XlsPath = filePath }

    member internal v.Body 
        with get() : string = v.XlsPath

    member v.ExportAsPdf(fitWidth:bool, quality:ExcelExportQuality, outFile:string) : PdfDoc = 
        withExcelApp <| fun app -> 
            try 
                let workbook : Excel.Workbook = app.Workbooks.Open(v.Body)
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
        let outFile:string = System.IO.Path.ChangeExtension(v.Body, "pdf")
        v.ExportAsPdf(fitWidth = fitWidth, quality = quality, outFile = outFile)

let excelDoc (path:string) : ExcelDoc = new ExcelDoc (filePath = path)