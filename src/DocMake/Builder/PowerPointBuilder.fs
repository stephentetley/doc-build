// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


module DocMake.Builder.PowerPointBuilder

open System.IO
open Microsoft.Office.Interop



open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis



type PowerPointBuild<'a> = BuildMonad<PowerPoint.Application, 'a>
type PowerPointPhantom = class end
type PowerPointDoc = Document<PowerPointPhantom>


let private initPowerPoint () : PowerPoint.Application = 
    let app = new PowerPoint.ApplicationClass() :> PowerPoint.Application
    app.Visible <- Microsoft.Office.Core.MsoTriState.msoTrue
    app

let private finalizePowerPoint (app:PowerPoint.Application) : unit = 
    app.Quit ()


let powerPointBuilderHook : BuilderHook<PowerPoint.Application> = 
    { InitializeResource = initPowerPoint
      FinalizeResource = finalizePowerPoint }






    