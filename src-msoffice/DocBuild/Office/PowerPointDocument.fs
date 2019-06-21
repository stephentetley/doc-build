// Copyright (c) Stephen Tetley 2018,2019
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

        interface IResourceFinalize with
            member x.RunFinalizer = 
                match x.PowerPointApplication with
                | null -> () 
                | app -> finalizePowerPoint app

        interface IPowerPointHandle with
            member x.PowerPointAppHandle = x

    and IPowerPointHandle =
        abstract PowerPointAppHandle : PowerPointHandle

    let execPowerPoint (mf: PowerPoint.Application -> DocMonad<'a, #IPowerPointHandle>) : DocMonad<'a, #IPowerPointHandle> = 
        docMonad { 
            let! userRes = getUserResources ()
            let powerPointHandle = userRes.PowerPointAppHandle
            return! mf powerPointHandle.PowerPointExe
        }

    // ************************************************************************
    // Export

    let private powerpointExportQuality (quality:PrintQuality) : PowerPoint.PpFixedFormatIntent = 
        match quality with
        | Screen -> PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentScreen
        | Print -> PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentPrint



    let exportPdfAs (outputRelName : string) 
                    (source : PowerPointDoc) : DocMonad<PdfDoc, #IPowerPointHandle> = 
        docMonad { 
            let! outputAbsPath = extendWorkingPath outputRelName
            let! sourcePath = getDocumentPath source
            let! pdfQuality = 
                asks (fun env -> env.PrintOrScreen) |>> powerpointExportQuality
            let! ans = 
                execPowerPoint <| fun app -> 
                    liftOperationResult "exportPdfAs" 
                        (fun _ -> powerPointExportAsPdf app pdfQuality sourcePath outputAbsPath)
            return! getPdfDoc outputAbsPath
        }

    /// Saves the file in the working directory.
    let exportPdf (source : PowerPointDoc) : DocMonad<PdfDoc, #IPowerPointHandle> = 
        docMonad { 
            let! sourceName = getDocumentFileName source
            let fileName = Path.ChangeExtension(sourceName, "pdf")
            return! exportPdfAs fileName source
        }
