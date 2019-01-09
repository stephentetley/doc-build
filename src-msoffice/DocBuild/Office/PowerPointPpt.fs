// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Office.PowerPointPpt


[<AutoOpen>]
module PowerPointPpt = 


    // Open at .Interop rather than .PowerPoint then the PowerPoint 
    // API has to be qualified
    open Microsoft.Office.Interop

    open DocBuild.Base.Document



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
        val private PowerPointDoc : Document

        new (filePath:string) = 
            { PowerPointDoc = new Document(filePath = filePath) }

        member x.ExportAsPdf( quality:PowerPointExportQuality
                            , outFile:string) : unit = 
            withPowerPointApp <| fun app -> 
                let srcFile = x.PowerPointDoc.ActiveFile
                try                     
                    let prez = app.Presentations.Open(srcFile)
                    prez.ExportAsFixedFormat (Path = outFile,
                                                FixedFormatType = PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF,
                                                Intent = powerpointExportQuality quality ) 
                    prez.Close()
                with
                | ex -> failwithf "PptToPdf - Some error occured for %s - '%s'" srcFile ex.Message


        member x.ExportAsPdf(quality:PowerPointExportQuality) : unit =
            // Don't make a temp file if we don't have to
            let srcFile = x.PowerPointDoc.ActiveFile
            let outFile:string = System.IO.Path.ChangeExtension(srcFile, "pdf")
            x.ExportAsPdf(quality= quality, outFile = outFile)

    let powerPointDoc (path:string) : PowerPointDoc = new PowerPointDoc (filePath = path)

