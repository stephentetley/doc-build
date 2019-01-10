// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Office


[<RequireQualifiedAccess>]
module PowerPointFile = 

    open System.IO

    // Open at .Interop rather than .PowerPoint then the PowerPoint 
    // API has to be qualified
    open Microsoft.Office.Interop

    open DocBuild.Base

    open DocBuild.Office
    open DocBuild.Office.Internal
    open DocBuild.Office.OfficeMonad

    // ************************************************************************
    // Export

    let private powerpointExportQuality (quality:PrintQuality) : PowerPoint.PpFixedFormatIntent = 
        match quality with
        | PqScreen -> PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentScreen
        | PqPrint -> PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentPrint



    let exportPdfAs (src:PowerPointFile) 
                    (quality:PrintQuality) 
                    (outputFile:string) : OfficeMonad<PdfFile> = 
        officeMonad { 
            let pdfQuality = powerpointExportQuality quality
            let! ans = 
                execPowerPoint <| fun app -> 
                    powerPointExportAsPdf app src.Path pdfQuality outputFile
            let! pdf = liftDocMonad (getPdfFile outputFile)
            return pdf
        }

    let exportPdf (src:PowerPointFile) (quality:PrintQuality) : OfficeMonad<PdfFile> = 
        let outputFile = Path.ChangeExtension(src.Path, "pdf")
        exportPdfAs src quality outputFile
