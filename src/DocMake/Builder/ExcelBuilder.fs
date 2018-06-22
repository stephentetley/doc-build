// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


module DocMake.Builder.ExcelBuilder

open System.IO
open Microsoft.Office.Interop

open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis


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

let withXlsxNamer (ma:BuildMonad<'res,'a>) : BuildMonad<'res,'a> = 
    withNameGen (sprintf "temp%03i.xlsx") ma



