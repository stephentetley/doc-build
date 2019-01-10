// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Office


[<AutoOpen>]
module Common = 
    

    open Microsoft.Office.Interop


    type PrintQuality = 
        | PqScreen
        | PqPrint

    //let internal wordPrintQuality (quality:PrintQuality) : Word.WdExportOptimizeFor = 
    //    match quality with
    //    | PqScreen -> Word.WdExportOptimizeFor.wdExportOptimizeForOnScreen
    //    | PqPrint -> Word.WdExportOptimizeFor.wdExportOptimizeForPrint
  

    let internal excelPrintQuality (quality:PrintQuality) : Excel.XlFixedFormatQuality = 
        match quality with
        | PqScreen -> Excel.XlFixedFormatQuality.xlQualityMinimum
        | PqPrint -> Excel.XlFixedFormatQuality.xlQualityStandard


    let internal powerPointPrintQuality (quality:PrintQuality) : PowerPoint.PpFixedFormatIntent = 
        match quality with
        | PqScreen -> PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentScreen
        | PqPrint -> PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentPrint

