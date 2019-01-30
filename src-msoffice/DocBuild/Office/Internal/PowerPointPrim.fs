// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Office.Internal


[<AutoOpen>]
module PowerPointPrim = 


    // Open at .Interop rather than .PowerPoint then the PowerPoint 
    // API has to be qualified
    open Microsoft.Office.Interop

    open DocBuild.Base.Document



    let withPowerPointApp (operation:PowerPoint.Application -> 'a) : 'a = 
        let app = new PowerPoint.ApplicationClass() :> PowerPoint.Application
        app.Visible <- Microsoft.Office.Core.MsoTriState.msoTrue
        let result = operation app
        app.Quit ()
        result




    // ****************************************************************************
    // Export to Pdf

    let powerPointExportAsPdf (app:PowerPoint.Application)
                              (quality:PowerPoint.PpFixedFormatIntent) 
                              (inputFile:string)
                              (outputFile:string ) : Result<unit,string> = 
        try
            withPowerPointApp <| fun app ->                 
                let prez = app.Presentations.Open(inputFile)
                prez.ExportAsFixedFormat( Path = outputFile
                                        , FixedFormatType = PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF
                                        , Intent = quality ) 
                prez.Close()
                Ok ()
        with
        | _ -> Error (sprintf "powerPointExportAsPdf failed '%s'" inputFile)
