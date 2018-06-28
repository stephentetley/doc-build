// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


module DocMake.Builder.ExcelHooks

open System.IO
open Microsoft.Office.Interop

open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis


type ExcelPhantom = class end
type ExcelDoc = Document<ExcelPhantom>



let internal initExcel () : Excel.Application = 
    let app = new Excel.ApplicationClass(Visible = true) :> Excel.Application
    app.DisplayAlerts <- false
    app.EnableEvents <- false
    app

let internal finalizeExcel (app:Excel.Application) : unit = 
        app.DisplayAlerts <- true
        app.EnableEvents <- true
        app.Quit ()


let withXlsxNamer (ma:BuildMonad<'res,'a>) : BuildMonad<'res,'a> = 
    withNameGen (sprintf "temp%03i.xlsx") ma



