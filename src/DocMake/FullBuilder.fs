// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


module DocMake.FullBuilder

open Microsoft.Office.Interop

open DocMake.Base.Common
open DocMake.Base.OfficeUtils
open DocMake.Builder.BuildMonad
open DocMake.Builder.Document
open DocMake.Builder.GhostscriptHooks
open DocMake.Builder.PdftkHooks
open DocMake.Tasks

type FullHandle (gs:GsHandle, pdftk:PdftkHandle) = 
    let gsHandle = gs
    let pdftkHandle = pdftk
    let mutable wordApp:Word.Application = null
    let mutable excelApp:Excel.Application = null
    let mutable powerPointApp:PowerPoint.Application = null

    member v.Ghostscript : GsHandle = gsHandle
    member v.Pdftk : PdftkHandle = pdftkHandle    

    member v.WordApp :Word.Application = 
        match wordApp with
        | null -> 
            let word1 = initWord ()
            wordApp <- word1
            word1
        | app -> app

    member v.ExcelApp :Excel.Application  = 
        match excelApp with
        | null -> 
            let excel1 = initExcel ()
            excelApp <- excel1
            excel1
        | app -> app

    member v.PowerPointApp :PowerPoint.Application  = 
        match powerPointApp with
        | null -> 
            let powerPoint1 = initPowerPoint ()
            powerPointApp <- powerPoint1
            powerPoint1
        | app -> app

    member v.RunFinalize () = 
        match wordApp with
        | null -> () 
        | app -> finalizeWord app
        match excelApp with
        | null -> ()
        | app -> finalizeExcel app
        match powerPointApp with
        | null -> ()
        | app -> finalizePowerPoint app


type FullBuild<'a> = BuildMonad<FullHandle,'a>

type FullBuildConfig  = 
    { GhostscriptPath: string
      PdftkPath: string } 

let runFullBuild (env:Env) (config:FullBuildConfig) (ma:FullBuild<'a>) : 'a = 
    let handle = new FullHandle({GhostscriptExePath = config.GhostscriptPath}, {PdftkExePath = config.PdftkPath })
    consoleRun env handle (fun (h:FullHandle) -> h.RunFinalize () ) ma


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

let pdfRotateAllCw (pdfDoc:PdfDoc) : FullBuild<PdfDoc> = 
    pdfRotateApi.PdfRotateAllCw pdfDoc

let pdfRotateAllCcw (pdfDoc:PdfDoc) : FullBuild<PdfDoc> = 
    pdfRotateApi.PdfRotateAllCcw pdfDoc


let rotationRange (startPage:int) (endPage:int) (orientation:PageOrientation) : PdfRotate.Rotation = 
    PdfRotate.rotationRange startPage endPage orientation

let rotationSinglePage (pageNum:int) (orientation:PageOrientation) : PdfRotate.Rotation = 
    PdfRotate.rotationSinglePage pageNum orientation

let rotationToEnd (startPage:int) (orientation:PageOrientation) : PdfRotate.Rotation = 
    PdfRotate.rotationToEnd startPage orientation

// *************************************
// Wraps DocMake.Tasks.DocPhotos

let docPhotos (opts:DocPhotos.DocPhotosOptions) (sourceDirectories:string list) : FullBuild<WordDoc> = 
    let api = DocPhotos.makeAPI (fun (h:FullHandle) -> h.WordApp)
    api.DocPhotos opts sourceDirectories
