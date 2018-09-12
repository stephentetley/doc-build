// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

[<RequireQualifiedAccess>]
module DocMake.Tasks.PptToPdf


// Open at .Interop rather than .PowerPoint then the PowerPoint API has to be qualified
open Microsoft.Office.Interop

open DocMake.Base.Common
open DocMake.Base.FakeLike
open DocMake.Base.OfficeUtils
open DocMake.Builder.BuildMonad
open DocMake.Builder.Document
open DocMake.Builder.Basis

   


let private process1 (inpath:string) (outpath:string) (quality:PrintQuality) (app:PowerPoint.Application) : unit = 
    try 
        let prez = app.Presentations.Open(inpath)
        prez.ExportAsFixedFormat (Path = outpath,
                                    FixedFormatType = PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF,
                                    Intent = powerpointPrintQuality quality) 
        prez.Close();
    with
    | ex -> printfn "PptToPdf - Some error occured for %s - '%s'" inpath ex.Message



/// Name is derived from the original name
/// Document is created in the working directory
let private pptToPdfImpl (getHandle:'res-> PowerPoint.Application) 
                (pptDoc:PowerPointDoc) : BuildMonad<'res,PdfDoc> =
    buildMonad { 
        let! (app:PowerPoint.Application) = asksU getHandle
        let! quality = asksEnv (fun s -> s.PrintQuality)
        match pptDoc.GetPath with
        | None -> return zeroDocument
        | Some pptPath -> 
            let name1 = System.IO.FileInfo(pptPath).Name
            let! path1 = askWorkingDirectory () |>> (fun cwd -> cwd </> name1)
            let outPath = System.IO.Path.ChangeExtension(path1, "pdf") 
            let _ =  process1 pptPath outPath quality app
            return (makeDocument outPath)
    }

type PptToPdfApi<'res> = 
    { PptToPdf : PowerPointDoc -> BuildMonad<'res, PdfDoc> }

let makeAPI (getHandle:'res-> PowerPoint.Application) : PptToPdfApi<'res> = 
    { PptToPdf = pptToPdfImpl getHandle }