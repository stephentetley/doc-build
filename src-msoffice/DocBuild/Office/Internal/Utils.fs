// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Office.Internal


[<AutoOpen>]
module Utils = 

    open Microsoft.Office.Interop

    
    let refobj (x:'a) : ref<obj> = ref (x :> obj)


    let initWord () : Word.Application = 
        let app = new Word.ApplicationClass (Visible = true) :> Word.Application
        app

    let finalizeWord (app:Word.Application) : unit = app.Quit ()

    let initExcel () : Excel.Application = 
        let app = new Excel.ApplicationClass(Visible = true) :> Excel.Application
        app.DisplayAlerts <- false
        app.EnableEvents <- false
        app

    let finalizeExcel (app:Excel.Application) : unit = 
        app.DisplayAlerts <- true
        app.EnableEvents <- true
        app.Quit ()


    let initPowerPoint () : PowerPoint.Application = 
        let app = new PowerPoint.ApplicationClass() :> PowerPoint.Application
        app.Visible <- Microsoft.Office.Core.MsoTriState.msoTrue
        app

    let finalizePowerPoint (app:PowerPoint.Application) : unit = 
        app.Quit ()


