// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

[<RequireQualifiedAccess>]
module DocMake.Tasks.XlsToPdf


// Open at .Interop rather than .Excel then the Excel API has to be qualified
open Microsoft.Office.Interop


open DocMake.Base.Common
open DocMake.Base.FakeLike
open DocMake.Base.OfficeUtils
open DocMake.Builder.BuildMonad
open DocMake.Builder.Document
open DocMake.Builder.Basis



let private process1 (inpath:string) (outpath:string) (quality:PrintQuality) (fitWidth:bool) (app:Excel.Application)  : unit = 
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



/// Name is derived from the original name
/// Document is created in the working directory
let private xlsToPdfImpl (getHandle:'res-> Excel.Application) (fitWidth:bool) (xlsDoc:ExcelDoc) : BuildMonad<'res, PdfDoc> =
    buildMonad { 
        let! (app:Excel.Application) = asksU getHandle
        let! quality = asksEnv (fun s -> s.PrintQuality)
        match xlsDoc.GetPath with
        | None -> return zeroDocument
        | Some xlsPath -> 
            let name1 = System.IO.FileInfo(xlsPath).Name
            let! path1 = askWorkingDirectory () |>> (fun cwd -> cwd </> name1)
            let outPath = System.IO.Path.ChangeExtension(path1, "pdf") 
            let _ =  process1 xlsPath outPath quality fitWidth app
            return (makeDocument outPath)
    }    
    

// The handle API is good that it loosely couples 
// resources (Excel, etc.) but it is too complicated.

type XlsToPdfApi<'res> = 
    { XlsToPdf : bool -> ExcelDoc -> BuildMonad<'res, PdfDoc> }

let makeAPI (getHandle:'res-> Excel.Application) : XlsToPdfApi<'res> = 
    { XlsToPdf = xlsToPdfImpl getHandle }