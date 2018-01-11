namespace DocMake.Base

open Microsoft.Office.Interop

open DocMake.Base.Common

module Office = 

    let refobj (x:'a) : ref<obj> = ref (x :> obj)

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


