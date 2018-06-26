// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

[<RequireQualifiedAccess>]
module DocMake.Tasks.PptToPdf


// Open at .Interop rather than .PowerPoint then the PowerPoint API has to be qualified
open Microsoft.Office.Interop


open DocMake.Base.Common
open DocMake.Base.OfficeUtils
open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis
open DocMake.Builder.PowerPointBuilder

   


let private process1 (inpath:string) (outpath:string) (quality:DocMakePrintQuality) (app:PowerPoint.Application) : unit = 
    try 
        let prez = app.Presentations.Open(inpath)
        prez.ExportAsFixedFormat (Path = outpath,
                                    FixedFormatType = PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF,
                                    Intent = powerpointPrintQuality quality) 
        prez.Close();
    with
    | ex -> printfn "PptToPdf - Some error occured for %s - '%s'" inpath ex.Message




let private pptToPdfImpl (getHandle:'res-> PowerPoint.Application) (pptDoc:PowerPointDoc) : BuildMonad<'res,PdfDoc> =
    buildMonad { 
        let! (app:PowerPoint.Application) = asksU getHandle
        let  outName = documentName <| documentChangeExtension "pdf" pptDoc
        let! outTemp = freshDocument () |>> documentChangeExtension "pdf"
        let! quality = asksEnv (fun s -> s.PrintQuality)
        let _ =  process1 pptDoc.DocumentPath outTemp.DocumentPath quality app
        let! final = renameTo outName outTemp
        return final
    }

type PptToPdf<'res> = 
    { pptToPdf : PowerPointDoc -> BuildMonad<'res, PdfDoc> }

let makeAPI (getHandle:'res-> PowerPoint.Application) : PptToPdf<'res> = 
    { pptToPdf = pptToPdfImpl getHandle }