// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

namespace DocBuild.Internal.ExcelUtils


[<AutoOpen>]
module ExcelUtils = 

    open Microsoft.Office.Interop

    open DocBuild.Base

    let internal withExcelApp (operation:Excel.Application -> 'a) : 'a = 
        let app = new Excel.ApplicationClass(Visible = true) :> Excel.Application
        app.DisplayAlerts <- false
        app.EnableEvents <- false
        let result = operation app
        app.DisplayAlerts <- true
        app.EnableEvents <- true
        app.Quit ()
        result

    // ****************************************************************************
    // Find/Replace

    let private sheetFindReplace (worksheet:Excel.Worksheet) (search:string, replace:string) : unit = 
        // TODO - I expect the fist line to work but it doesn't
        // let allCells : Excel.Range = worksheet.Cells
        let allCells : Excel.Range = worksheet.Range("A1")
    
        // Maybe this has to be iterated until false...
        allCells.Replace( What = search, 
                          Replacement = replace,
                          LookAt = Excel.XlLookAt.xlWhole,
                          SearchOrder = Excel.XlSearchOrder.xlByColumns,
                          MatchCase = true,
                          SearchFormat = false,
                          ReplaceFormat = false ) |> ignore

    
    let private sheetFindReplaceList (worksheet:Excel.Worksheet) (searches:SearchList) : unit =                      
        List.iter (sheetFindReplace worksheet) searches


    let workbookFindReplace (workbook:Excel.Workbook) (searches:SearchList) : unit = 
        workbook.Worksheets 
            |> Seq.cast<Excel.Worksheet> 
            |> Seq.iter (fun (sheet:Excel.Worksheet) -> sheetFindReplaceList sheet searches)


    let excelFindReplace (app:Excel.Application) 
                         (inpath:string) 
                         (outpath:option<string>) 
                         (searches:SearchList) : unit = 
        let workbook : Excel.Workbook = app.Workbooks.Open(inpath)
        workbookFindReplace workbook searches
        try 
            match outpath with 
            | None -> workbook.Save()
            | Some filename -> 
                printfn "Outpath: %s" filename
                workbook.SaveAs (Filename = filename)
        finally 
            workbook.Close (SaveChanges = false)


