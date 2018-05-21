module DocMake.Lib.XlsToPdf


// Open at .Interop rather than .Excel then the Excel API has to be qualified
open Microsoft.Office.Interop


open DocMake.Base.Common
open DocMake.Base.OfficeUtils
open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis
open DocMake.Builder.Builders


let private getOutputName (xlsDoc:ExcelDoc) : ExcelBuild<string> =
    executeIO <| fun () -> 
        System.IO.Path.ChangeExtension(xlsDoc.DocumentPath, "pdf")
    


let private process1 (inpath:string) (outpath:string) (quality:DocMakePrintQuality) (fitWidth:bool) (app:Excel.Application)  : unit = 
    try 
        let workbook : Excel.Workbook = app.Workbooks.Open(inpath)
        if fitWidth then 
            workbook.Sheets 
                |> Seq.cast<Excel.Worksheet>
                |> Seq.iter (fun (sheet:Excel.Worksheet) -> 
                    sheet.PageSetup.Zoom <- false
                    sheet.PageSetup.FitToPagesWide <- 1)
        else ()

        workbook.ExportAsFixedFormat (Type=Excel.XlFixedFormatType.xlTypePDF,
                                         Filename=outpath,
                                         IncludeDocProperties=true,
                                         Quality = excelPrintQuality quality
                                         )
        workbook.Close (SaveChanges = false)
    with
    | ex -> printfn "%s" ex.Message





let xlsToPdf (fitWidth:bool) (xlsDoc:ExcelDoc) : ExcelBuild<PdfDoc> =
    buildMonad { 
        let! (app:Excel.Application) = askU ()
        let! outPath = getOutputName xlsDoc
        let! quality = asksEnv (fun s -> s.PrintQuality)
        let _ =  process1 xlsDoc.DocumentPath outPath quality fitWidth app
        return (makeDocument outPath)
    }