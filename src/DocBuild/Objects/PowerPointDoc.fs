// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


namespace DocBuild


[<AutoOpen>]
module PowerPointDoc = 


    // Open at .Interop rather than .PowerPoint then the PowerPoint 
    // API has to be qualified
    open Microsoft.Office.Interop

    open DocBuild.Internal.CommonUtils
    open DocBuild


    let private withPowerPointApp (operation:PowerPoint.Application -> 'a) : 'a = 
        let app = new PowerPoint.ApplicationClass() :> PowerPoint.Application
        app.Visible <- Microsoft.Office.Core.MsoTriState.msoTrue
        let result = operation app
        app.Quit ()
        result


    type PowerPointExportQuality = 
        | PowerPointForScreen
        | PowerPointForPrint


    let private powerpointExportQuality (quality:PowerPointExportQuality) : PowerPoint.PpFixedFormatIntent = 
        match quality with
        | PowerPointForScreen -> PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentScreen
        | PowerPointForPrint -> PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentPrint


    type PowerPointDoc = 
        val private PptPath : string

        new (filePath:string) = 
            { PptPath = filePath }

        member internal v.Body 
            with get() : string = v.PptPath

        member v.ExportAsPdf( quality:PowerPointExportQuality
                            , outFile:string) : Document = 
            withPowerPointApp <| fun app -> 
                try 
                    let prez = app.Presentations.Open(v.Body)
                    prez.ExportAsFixedFormat (Path = outFile,
                                                FixedFormatType = PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF,
                                                Intent = powerpointExportQuality quality ) 
                    prez.Close()
                    pdfDoc outFile
                with
                | ex -> failwithf "PptToPdf - Some error occured for %s - '%s'" v.Body ex.Message


        member v.ExportAsPdf(quality:PowerPointExportQuality) : Document =
            let outFile:string = System.IO.Path.ChangeExtension(v.Body, "pdf")
            v.ExportAsPdf(quality= quality, outFile = outFile)

    let powerPointDoc (path:string) : PowerPointDoc = new PowerPointDoc (filePath = path)

