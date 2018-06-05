﻿// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

[<RequireQualifiedAccess>]
module DocMake.Lib.XlsToPdf


// Open at .Interop rather than .Excel then the Excel API has to be qualified
open Microsoft.Office.Interop


open DocMake.Base.Common
open DocMake.Base.OfficeUtils
open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis
open DocMake.Builder.ExcelBuilder



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
        let! outPath = freshDocument () |>> documentChangeExtension "pdf"
        let! quality = asksEnv (fun s -> s.PrintQuality)
        let _ =  process1 xlsDoc.DocumentPath outPath.DocumentPath quality fitWidth app
        return outPath
    }    
    

type XlsToPdf<'res> = 
    { xlsToPdf : bool -> ExcelDoc -> BuildMonad<'res, PdfDoc> }

let makeAPI (getHandle:'res-> Excel.Application) : XlsToPdf<'res> = 
    { xlsToPdf = xlsToPdfImpl getHandle }