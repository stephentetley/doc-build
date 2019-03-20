﻿// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Office


[<RequireQualifiedAccess>]
module PowerPointDocument = 

    open System.IO

    // Open at .Interop rather than .PowerPoint then the PowerPoint 
    // API has to be qualified
    open Microsoft.Office.Interop

    open DocBuild.Base
    open DocBuild.Base.DocMonad
    open DocBuild.Base.DocMonadOperators
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

        interface ResourceFinalize with
            member x.RunFinalizer = 
                match x.PowerPointApplication with
                | null -> () 
                | app -> finalizePowerPoint app

        interface HasPowerPointHandle with
            member x.PowerPointAppHandle = x

    and HasPowerPointHandle =
        abstract PowerPointAppHandle : PowerPointHandle

    let execPowerPoint (mf: PowerPoint.Application -> DocMonad<#HasPowerPointHandle,'a>) : DocMonad<#HasPowerPointHandle,'a> = 
        docMonad { 
            let! userRes = askUserResources ()
            let powerPointHandle = userRes.PowerPointAppHandle
            return! mf powerPointHandle.PowerPointExe
        }

    // ************************************************************************
    // Export

    let private powerpointExportQuality (quality:PrintQuality) : PowerPoint.PpFixedFormatIntent = 
        match quality with
        | Screen -> PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentScreen
        | Print -> PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentPrint



    let exportPdfAs (outputAbsPath:string) 
                    (src:PowerPointDoc) : DocMonad<#HasPowerPointHandle,PdfDoc> = 
        docMonad { 
            do! assertIsWorkingPath outputAbsPath
            let! pdfQuality = 
                asks (fun env -> env.PrintOrScreen) |>> powerpointExportQuality
            let! ans = 
                execPowerPoint <| fun app -> 
                    liftResult (powerPointExportAsPdf app pdfQuality src.LocalPath outputAbsPath)
            return! workingPdfDoc outputAbsPath
        }

    /// Saves the file in the working directory.
    let exportPdf (src:PowerPointDoc) : DocMonad<#HasPowerPointHandle,PdfDoc> = 
        docMonad { 
            let! path1 = extendWorkingPath src.FileName
            let outputAbsPath = Path.ChangeExtension(path1, "pdf")
            return! exportPdfAs outputAbsPath src
        }
