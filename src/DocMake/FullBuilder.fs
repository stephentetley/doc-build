module DocMake.FullBuilder


open Microsoft.Office.Interop

open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis
open DocMake.Builder.WordBuilder
open DocMake.Builder.ExcelBuilder
open DocMake.Builder.PowerPointBuilder
open DocMake.Builder.GhostscriptBuilder
open DocMake.Builder.PdftkBuilder
open DocMake.Lib


/// Note - we probably need to look at "by need" creation of Excel, 
/// PowerPoint, etc. instances.
type FullHandle = 
    { WordApp : Word.Application
      ExcelApp : Excel.Application 
      PowerPointApp : PowerPoint.Application
      Ghostscript: GsHandle
      Pdftk: PdftkHandle
      }

type FullBuild<'a> = BuildMonad<FullHandle,'a>

let private initFullBuilder (gsPath:string) (pdftkPath:string) : FullHandle = 
    { WordApp = wordBuilderHook.InitializeResource ()
      ExcelApp = excelBuilderHook.InitializeResource () 
      PowerPointApp = powerPointBuilderHook.InitializeResource ()
      Ghostscript = ghostsciptBuilderHook(gsPath).InitializeResource ()
      Pdftk = pdftkBuilderHook(pdftkPath).InitializeResource ()
      }

let private finalizeFullBuilder (handle:FullHandle) : unit = 
    wordBuilderHook.FinalizeResource handle.WordApp
    excelBuilderHook.FinalizeResource handle.ExcelApp
    powerPointBuilderHook.FinalizeResource handle.PowerPointApp
    // do nothing for Ghostscript
    // do nothing for Pdftk


let fullBuilderHooks (gsPath:string) (pdftkPath:string) : BuilderHooks<FullHandle> = 
    { InitializeResource  = fun _ -> initFullBuilder gsPath pdftkPath
      FinalizeResource = finalizeFullBuilder }


let d1 : DocMake.Lib.DocFindReplace.DocFindReplace<FullHandle> = 
    DocMake.Lib.DocFindReplace.makeAPI (fun (h:FullHandle) -> h.WordApp)

let docFindReplace = d1.docFindReplace
let getTemplate = d1.getTemplate

let docToPdf (wordDoc:WordDoc) : FullBuild<PdfDoc> = 
    let api = DocToPdf.makeAPI (fun (h:FullHandle) -> h.WordApp)
    api.docToPdf wordDoc 

let xlsToPdf (fitPage:bool) (xlsDoc:ExcelDoc) : FullBuild<PdfDoc> = 
    let api = XlsToPdf.makeAPI (fun (h:FullHandle) -> h.ExcelApp)
    api.xlsToPdf fitPage xlsDoc

let pptToPdf (pptDoc:PowerPointDoc) : FullBuild<PdfDoc> = 
    let api = PptToPdf.makeAPI (fun (h:FullHandle) -> h.PowerPointApp)
    api.pptToPdf pptDoc


let docPhotos (opts:DocPhotos.DocPhotosOptions) (sourceDirectories:string list) : FullBuild<WordDoc> = 
    let api = DocPhotos.makeAPI (fun (h:FullHandle) -> h.WordApp)
    api.docPhotos opts sourceDirectories

let pdfConcat (inputFiles:PdfDoc list) : FullBuild<PdfDoc> = 
    let api = PdfConcat.makeAPI (fun (h:FullHandle) -> h.Ghostscript)
    api.pdfConcat inputFiles

let pdfRotate (rotations: PdfRotate.PageRotation list) (pdfDoc:PdfDoc) : FullBuild<PdfDoc> = 
    let api = PdfRotate.makeAPI (fun (h:FullHandle) -> h.Pdftk)
    api.pdfRotate rotations pdfDoc

