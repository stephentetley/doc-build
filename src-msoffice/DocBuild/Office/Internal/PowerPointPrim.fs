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


