module DocMake.Base.OfficeUtils

open Microsoft.Office.Interop

open DocMake.Base.Common


let refobj (x:'a) : ref<obj> = ref (x :> obj)

let wordPrintQuality (quality:DocMakePrintQuality) : Word.WdExportOptimizeFor = 
    match quality with
    | PqScreen -> Word.WdExportOptimizeFor.wdExportOptimizeForOnScreen
    | PqPrint -> Word.WdExportOptimizeFor.wdExportOptimizeForPrint
  

let excelPrintQuality (quality:DocMakePrintQuality) : Excel.XlFixedFormatQuality = 
    match quality with
    | PqScreen -> Excel.XlFixedFormatQuality.xlQualityMinimum
    | PqPrint -> Excel.XlFixedFormatQuality.xlQualityStandard

let powerpointPrintQuality (quality:DocMakePrintQuality) : PowerPoint.PpFixedFormatIntent = 
    match quality with
    | PqScreen -> PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentScreen
    | PqPrint -> PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentPrint


