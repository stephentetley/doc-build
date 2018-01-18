// Potentially these helper modules in DocMake.Base need 
// longer / better names.
// If the file is missed from the #load directives in an .fsx script
// the warning is too generic:
// "The namespace 'Office' is not defined."

module DocMake.Base.Office

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


