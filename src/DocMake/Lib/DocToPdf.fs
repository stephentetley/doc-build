// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

[<RequireQualifiedAccess>]
module DocMake.Lib.DocToPdf


// Open at .Interop rather than .Word then the Word API has to be qualified
open Microsoft.Office.Interop


open DocMake.Base.Common
open DocMake.Base.OfficeUtils
open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis
open DocMake.Builder.WordBuilder

    

let private process1 (inpath:string) (outpath:string) (quality:DocMakePrintQuality) (app:Word.Application) : unit = 
    try 
        let doc = app.Documents.Open(FileName = refobj inpath)
        doc.ExportAsFixedFormat (OutputFileName = outpath, 
                                  ExportFormat = Word.WdExportFormat.wdExportFormatPDF,
                                  OptimizeFor = wordPrintQuality quality)
        doc.Close (SaveChanges = refobj false)
    with
    | ex -> printfn "Some error occured - %s - %s" inpath ex.Message



let private docToPdfImpl (getHandle:'res-> Word.Application) (wordDoc:WordDoc) : BuildMonad<'res,PdfDoc> =
    buildMonad { 
        let! (app:Word.Application) = asksU getHandle
        let! outPath = freshDocument () |>> documentChangeExtension "pdf"
        let! quality = asksEnv (fun s -> s.PrintQuality)
        let _ =  process1 wordDoc.DocumentPath outPath.DocumentPath quality app
        return outPath
    }

    
type DocToPdf<'res> = 
    { docToPdf : WordDoc -> BuildMonad<'res, PdfDoc> }

let makeAPI (getHandle:'res-> Word.Application) : DocToPdf<'res> = 
    { docToPdf = docToPdfImpl getHandle }
