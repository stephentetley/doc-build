// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


module DocMake.Builder.ExcelBuilder

open System.IO
open Microsoft.Office.Interop



open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis


type ExcelBuild<'a> = BuildMonad<Excel.Application, 'a>
type ExcelPhantom = class end
type ExcelDoc = Document<ExcelPhantom>

let castToExcelDoc (doc:Document<'a>) : ExcelDoc = castDocument doc


let private initExcel () : Excel.Application = 
    let app = new Excel.ApplicationClass(Visible = true) :> Excel.Application
    app.DisplayAlerts <- false
    app.EnableEvents <- false
    app

let private finalizeExcel (app:Excel.Application) : unit = 
        app.DisplayAlerts <- true
        app.EnableEvents <- true
        app.Quit ()


let excelBuilderHook : BuilderHooks<Excel.Application> = 
    { InitializeResource = initExcel    
      FinalizeResource = finalizeExcel }





// TODO - Remove - run as a global single instance...
let execExcelBuild (ma:ExcelBuild<'a>) : BuildMonad<'res,'a> = 
    let app = new Excel.ApplicationClass(Visible = true) :> Excel.Application
    app.DisplayAlerts <- false
    app.EnableEvents <- false
    let namer:int -> string = fun i -> sprintf "temp%03i.docx" i
    let finalizer (oApp:Excel.Application) = 
        oApp.DisplayAlerts <- true
        oApp.EnableEvents <- true
        oApp.Quit ()
    withUserHandle app finalizer (withNameGen namer ma)


    