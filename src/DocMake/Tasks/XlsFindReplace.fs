// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

module DocMake.Tasks.XlsFindReplace

open System.IO
open Microsoft.Office.Interop

open Fake
open Fake.Core

open DocMake.Base.Common
open DocMake.Base.JsonUtils
open DocMake.Base.OfficeUtils


// The representation of matches is a SearchList (rather than JSON).
// The client can then work out weather or not it is serialized.
// DocFindReplace should be changed accordingly

[<CLIMutable>]
type XlsFindReplaceParams = 
    { TemplateFile: string
      Matches: SearchList
      OutputFile: string }


let XlsFindReplaceDefaults = 
    { TemplateFile = @""
      Matches = []
      OutputFile = @"findreplace.xlsx" }




let private sheetReplaces1 (worksheet:Excel.Worksheet) (search:string, replace:string) : unit = 
    // TODO - I expect the fist line to work but it doesn't
    // let allCells : Excel.Range = worksheet.Cells
    let allCells : Excel.Range = worksheet.Range("A1")
    // Maybe this has to be iterated until false...
    ignore <| 
        allCells.Replace(What = search, 
                            Replacement = replace,
                            LookAt = Excel.XlLookAt.xlWhole,
                            SearchOrder = Excel.XlSearchOrder.xlByColumns,
                            MatchCase = true,
                            SearchFormat = false,
                            ReplaceFormat = false )

    
let private sheetReplaces (worksheet:Excel.Worksheet) (searches:SearchList) : unit =                      
    List.iter (sheetReplaces1 worksheet) searches


let private replaces (workbook:Excel.Workbook) (searches:SearchList) : unit = 
    workbook.Worksheets 
        |> Seq.cast<Excel.Worksheet> 
        |> Seq.iter (fun (sheet:Excel.Worksheet) -> sheetReplaces sheet searches)


let private process1 (app:Excel.Application) (inpath:string) (outpath:string) (ss:SearchList) = 
    // This should be wrapped in try...
    let workbook : Excel.Workbook = app.Workbooks.Open(inpath)
    try 
        replaces workbook ss
        printfn "Outpath: %s" outpath
        workbook.SaveAs (Filename = outpath)
    finally 
        workbook.Close (SaveChanges = false)


let XlsFindReplace (setXlsFindReplaceParams: XlsFindReplaceParams -> XlsFindReplaceParams) : unit =
    let options = XlsFindReplaceDefaults |> setXlsFindReplaceParams
    if File.Exists(options.TemplateFile) then
        let app = new Excel.ApplicationClass(Visible = true)
        app.DisplayAlerts <- false
        app.EnableEvents <- false
        try 
            process1 app options.TemplateFile options.OutputFile options.Matches
            app.DisplayAlerts <- true
            app.EnableEvents <- true
        finally 
            app.Quit ()
    else 
        Trace.traceError <| sprintf "XlsFindReplace --- missing template file '%s'" options.TemplateFile
        failwith "XlsFindReplace --- missing template file"



