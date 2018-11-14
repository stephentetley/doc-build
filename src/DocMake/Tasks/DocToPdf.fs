// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

[<RequireQualifiedAccess>]
module DocMake.Tasks.DocToPdf


// Open at .Interop rather than .Word then the Word API has to be qualified
open Microsoft.Office.Interop


open DocMake.Base.Common
open DocMake.Base.FakeLike
open DocMake.Base.OfficeUtils
open DocMake.Builder.BuildMonad
open DocMake.Builder.Document
open DocMake.Builder.Basis
open DocMake.Builder.WordRunner
open System.IO
    

let private process1 (app:Word.Application) (inpath:string) (outpath:string) (quality:PrintQuality)  : unit = 
    try 
        let doc = app.Documents.Open(FileName = refobj inpath)
        doc.ExportAsFixedFormat (OutputFileName = outpath, 
                                  ExportFormat = Word.WdExportFormat.wdExportFormatPDF,
                                  OptimizeFor = wordPrintQuality quality)
        doc.Close (SaveChanges = refobj false)
    with
    | ex -> printfn "Some error occured - %s - %s" inpath ex.Message


/// Name is derived from the original name
/// Document is created in the working directory
/// WARNING - the naming startegy is  bad - if the file is "temp0??.docx"
/// it can potentially overwrite other working files.
let private docToPdfImpl (getHandle:'res-> Word.Application) (wordDoc:WordDoc) : BuildMonad<'res,PdfDoc> =
    buildMonad { 
        let! (app:Word.Application) = asksU getHandle
        let! quality = asksEnv (fun s -> s.PrintQuality)
        match wordDoc.GetPath with
        | None -> return zeroDocument
        | Some docPath ->
            let name1 = System.IO.FileInfo(docPath).Name
            let! path1 = askWorkingDirectory () |>> (fun cwd -> cwd </> name1)
            let outPath = System.IO.Path.ChangeExtension(path1, "pdf") 
            let _ = process1 app docPath outPath quality 
            return (makeDocument outPath)
    }

    
type DocToPdfApi<'res> = 
    { DocToPdf : WordDoc -> BuildMonad<'res, PdfDoc> }

let makeAPI (getHandle:'res-> Word.Application) : DocToPdfApi<'res> = 
    { DocToPdf = docToPdfImpl getHandle }

// ****************************************************************************

/// New API

let docToPdf (doc:WordDoc) (outputPath:string) : WordRunner<PdfDoc> = 
    match doc.GetPath with
    | None -> liftBM <| throwError "invalid Doc path"
    | Some path-> 
        wordRunner { 
            let! quality = liftBM <| asksEnv (fun s -> s.PrintQuality)
            let! _ = 
                wordExec <| fun wordApp -> 
                    process1 wordApp path outputPath quality 
            return (makeDocument outputPath)
        }


let runDocToPdf (doc:WordDoc) (outputPath:string) : BuildMonad<unit, PdfDoc> = 
    wordRun (docToPdf doc outputPath)