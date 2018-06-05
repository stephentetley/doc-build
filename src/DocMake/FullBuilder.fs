module DocMake.FullBuilder


open Microsoft.Office.Interop

open DocMake.Builder.BuildMonad
open DocMake.Builder.WordBuilder
open DocMake.Builder.ExcelBuilder
open DocMake.Builder.PowerPointBuilder
open DocMake.Builder.GhostscriptBuilder
open DocMake.Builder.PdftkBuilder
open DocMake.Lib


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

let docToPdf = 
    let api = DocMake.Lib.DocToPdf.makeAPI (fun (h:FullHandle) -> h.WordApp)
    api.docToPdf

let xlsToPdf = 
    let api = DocMake.Lib.XlsToPdf.makeAPI (fun (h:FullHandle) -> h.ExcelApp)
    api.xlsToPdf

let docPhotos = 
    let api = DocMake.Lib.DocPhotos.makeAPI (fun (h:FullHandle) -> h.WordApp)
    api.docPhotos

let pdfConcat = 
    let api = DocMake.Lib.PdfConcat.makeAPI (fun (h:FullHandle) -> h.Ghostscript)
    api.pdfConcat
