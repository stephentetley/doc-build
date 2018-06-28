module DocMake.FullBuilder


open Microsoft.Office.Interop


open DocMake.Base.Common
open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis
open DocMake.Builder.WordHooks
open DocMake.Builder.ExcelHooks
open DocMake.Builder.PowerPointHooks
open DocMake.Builder.GhostscriptHooks
open DocMake.Builder.PdftkHooks
open DocMake.Tasks


/// Note - we need to look at "by need" creation of Excel, Word 
/// and PowerPoint instances.
type FullHandle = 
    { WordApp: Word.Application
      ExcelApp: Excel.Application 
      PowerPointApp: PowerPoint.Application
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


// *************************************
// Wraps DocMake.Tasks.DocFindReplace


/// DocFindReplace Api has more than one entry point...
let private docFindReplaceApi : DocFindReplace.DocFindReplace<FullHandle> = 
    DocFindReplace.makeAPI (fun (h:FullHandle) -> h.WordApp)

let getTemplateDoc (docPath:string) : FullBuild<WordDoc> = 
    docFindReplaceApi.GetTemplateDoc docPath


let docFindReplace (searchList:SearchList) (template:WordDoc) : FullBuild<WordDoc> = 
    docFindReplaceApi.DocFindReplace searchList template


// *************************************
// Wraps DocMake.Tasks.XlsFindReplace

/// XlsFindReplace Api has more than one entry point...
let private xlsFindReplaceApi : XlsFindReplace.XlsFindReplace<FullHandle> = 
    XlsFindReplace.makeAPI (fun (h:FullHandle) -> h.ExcelApp)

let getTemplateXls (xlsPath:string) : FullBuild<ExcelDoc> = 
    xlsFindReplaceApi.GetTemplateXls xlsPath


let xlsFindReplace (searchList:SearchList) (template:ExcelDoc) : FullBuild<ExcelDoc> = 
    xlsFindReplaceApi.XlsFindReplace searchList template

    

// *************************************
// Wraps DocMake.Tasks.DocToPdf

let docToPdf (wordDoc:WordDoc) : FullBuild<PdfDoc> = 
    let api = DocToPdf.makeAPI (fun (h:FullHandle) -> h.WordApp)
    api.DocToPdf wordDoc 


// *************************************
// Wraps DocMake.Tasks.XlsToPdf

let xlsToPdf (fitPage:bool) (xlsDoc:ExcelDoc) : FullBuild<PdfDoc> = 
    let api = XlsToPdf.makeAPI (fun (h:FullHandle) -> h.ExcelApp)
    api.XlsToPdf fitPage xlsDoc

    
// *************************************
// Wraps DocMake.Tasks.PptToPdf

let pptToPdf (pptDoc:PowerPointDoc) : FullBuild<PdfDoc> = 
    let api = PptToPdf.makeAPI (fun (h:FullHandle) -> h.PowerPointApp)
    api.PptToPdf pptDoc


// *************************************
// Wraps DocMake.Tasks.PdfConcat

let pdfConcat (inputFiles:PdfDoc list) : FullBuild<PdfDoc> = 
    let api = PdfConcat.makeAPI (fun (h:FullHandle) -> h.Ghostscript)
    api.PdfConcat inputFiles


// *************************************
// Wraps DocMake.Tasks.PdfRotate


/// XlsFindReplace Api has more than one entry point...
let private pdfRotateApi : PdfRotate.PdfRotate<FullHandle> = 
    PdfRotate.makeAPI (fun (h:FullHandle) -> h.Pdftk)


let pdfRotateEmbed (rotations: PdfRotate.Rotation list) (pdfDoc:PdfDoc) : FullBuild<PdfDoc> = 
    pdfRotateApi.PdfRotateEmbed rotations pdfDoc

let pdfRotateExtract (rotations: PdfRotate.Rotation list) (pdfDoc:PdfDoc) : FullBuild<PdfDoc> = 
    pdfRotateApi.PdfRotateExtract rotations pdfDoc

let pdfRotateAll (orientation: PageOrientation) (pdfDoc:PdfDoc) : FullBuild<PdfDoc> = 
    pdfRotateApi.PdfRotateAll orientation pdfDoc


// *************************************
// Wraps DocMake.Tasks.DocPhotos

let docPhotos (opts:DocPhotos.DocPhotosOptions) (sourceDirectories:string list) : FullBuild<WordDoc> = 
    let api = DocPhotos.makeAPI (fun (h:FullHandle) -> h.WordApp)
    api.DocPhotos opts sourceDirectories
