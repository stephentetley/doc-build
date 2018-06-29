﻿// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


module DocMake.Builder.PowerPointHooks

open System.IO
open Microsoft.Office.Interop

open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis



type PowerPointPhantom = class end
type PowerPointDoc = Document<PowerPointPhantom>


let internal initPowerPoint () : PowerPoint.Application = 
    let app = new PowerPoint.ApplicationClass() :> PowerPoint.Application
    app.Visible <- Microsoft.Office.Core.MsoTriState.msoTrue
    app

let internal finalizePowerPoint (app:PowerPoint.Application) : unit = 
    app.Quit ()




    