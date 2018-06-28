// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

[<RequireQualifiedAccess>]
module DocMake.Tasks.XlsToPdf


// Open at .Interop rather than .Excel then the Excel API has to be qualified
open Microsoft.Office.Interop


open DocMake.Base.Common
open DocMake.Base.OfficeUtils
open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis
open DocMake.Builder.ExcelHooks



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





let private xlsToPdfImpl (getHandle:'res-> Excel.Application) (fitWidth:bool) (xlsDoc:ExcelDoc) : BuildMonad<'res, PdfDoc> =
    buildMonad { 
        let! (app:Excel.Application) = asksU getHandle
        let  outName = documentName <| documentChangeExtension "pdf" xlsDoc
        let! outTemp = freshDocument () |>> documentChangeExtension "pdf"
        let! quality = asksEnv (fun s -> s.PrintQuality)
        let _ =  process1 xlsDoc.DocumentPath outTemp.DocumentPath quality fitWidth app
        let! final = renameTo outName outTemp
        return final
    }    
    

type XlsToPdf<'res> = 
    { XlsToPdf : bool -> ExcelDoc -> BuildMonad<'res, PdfDoc> }

let makeAPI (getHandle:'res-> Excel.Application) : XlsToPdf<'res> = 
    { XlsToPdf = xlsToPdfImpl getHandle }