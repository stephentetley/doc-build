// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

module DocMake.Base.OfficeUtils

open Microsoft.Office.Interop

open DocMake.Base.Common


let refobj (x:'a) : ref<obj> = ref (x :> obj)


let internal initWord () : Word.Application = 
    let app = new Word.ApplicationClass (Visible = true) :> Word.Application
    app

let internal finalizeWord (app:Word.Application) : unit = app.Quit ()


let internal initExcel () : Excel.Application = 
    let app = new Excel.ApplicationClass(Visible = true) :> Excel.Application
    app.DisplayAlerts <- false
    app.EnableEvents <- false
    app

let internal finalizeExcel (app:Excel.Application) : unit = 
        app.DisplayAlerts <- true
        app.EnableEvents <- true
        app.Quit ()



let internal initPowerPoint () : PowerPoint.Application = 
    let app = new PowerPoint.ApplicationClass() :> PowerPoint.Application
    app.Visible <- Microsoft.Office.Core.MsoTriState.msoTrue
    app

let internal finalizePowerPoint (app:PowerPoint.Application) : unit = 
    app.Quit ()





let wordPrintQuality (quality:PrintQuality) : Word.WdExportOptimizeFor = 
    match quality with
    | PqScreen -> Word.WdExportOptimizeFor.wdExportOptimizeForOnScreen
    | PqPrint -> Word.WdExportOptimizeFor.wdExportOptimizeForPrint
  

let excelPrintQuality (quality:PrintQuality) : Excel.XlFixedFormatQuality = 
    match quality with
    | PqScreen -> Excel.XlFixedFormatQuality.xlQualityMinimum
    | PqPrint -> Excel.XlFixedFormatQuality.xlQualityStandard

let powerpointPrintQuality (quality:PrintQuality) : PowerPoint.PpFixedFormatIntent = 
    match quality with
    | PqScreen -> PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentScreen
    | PqPrint -> PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentPrint


