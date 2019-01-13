﻿// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Office


[<RequireQualifiedAccess>]
module PowerPointFile = 

    open System.IO

    // Open at .Interop rather than .PowerPoint then the PowerPoint 
    // API has to be qualified
    open Microsoft.Office.Interop

    open DocBuild.Base
    open DocBuild.Base.DocMonad
    open DocBuild.Office.Internal

    type PowerPointHandle = 
        val mutable private PowerPointApplication : PowerPoint.Application 

        new () = 
            { PowerPointApplication = null }

        /// Opens a handle as needed.
        member x.PowerPointExe : PowerPoint.Application = 
            match x.PowerPointApplication with
            | null -> 
                let powerPoint1 = initPowerPoint ()
                x.PowerPointApplication <- powerPoint1
                powerPoint1
            | app -> app

        member x.RunFinalizer () = 
            match x.PowerPointApplication with
            | null -> () 
            | app -> finalizePowerPoint app

    type HasPowerPointHandle =
        abstract PowerPointAppHandle : PowerPointHandle

    let execPowerPoint<'res when 'res :> HasPowerPointHandle> 
                      (mf: PowerPoint.Application -> DocMonad<'res,'a>) : DocMonad<'res,'a> = 
        docMonad { 
            let! userRes = askUserResources ()
            let powerPointHandle = userRes.PowerPointAppHandle
            let! ans = mf powerPointHandle.PowerPointExe
            return ans
        }

    // ************************************************************************
    // Export

    let private powerpointExportQuality (quality:PrintQuality) : PowerPoint.PpFixedFormatIntent = 
        match quality with
        | PqScreen -> PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentScreen
        | PqPrint -> PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentPrint



    let exportPdfAs (src:PowerPointFile) 
                    (quality:PrintQuality) 
                    (outputFile:string) : DocMonad<#HasPowerPointHandle,PdfFile> = 
        docMonad { 
            let pdfQuality = powerpointExportQuality quality
            let! ans = 
                execPowerPoint <| fun app -> 
                    liftResult (powerPointExportAsPdf app src.Path pdfQuality outputFile)
            let! pdf = getPdfFile outputFile
            return pdf
        }

    let exportPdf (src:PowerPointFile) (quality:PrintQuality) : DocMonad<#HasPowerPointHandle,PdfFile> = 
        let outputFile = Path.ChangeExtension(src.Path, "pdf")
        exportPdfAs src quality outputFile
