// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause

namespace DocBuild.Office.Internal



[<AutoOpen>]
module ExcelPrim = 

    open Microsoft.Office.Interop

    open DocBuild.Base.Common

    let withExcelApp (operation:Excel.Application -> 'a) : 'a = 
        let app = new Excel.ApplicationClass(Visible = true) :> Excel.Application
        app.DisplayAlerts <- false
        app.EnableEvents <- false
        let result = operation app
        app.DisplayAlerts <- true
        app.EnableEvents <- true
        app.Quit ()
        result

    // ****************************************************************************
    // Export to Pdf

    let excelExportAsPdf (app:Excel.Application) 
                         (inputFile:string)
                         (fitWidth:bool)
                         (quality:Excel.XlFixedFormatQuality)
                         (outputFile:string ) : Result<unit,string> = 
        try
            withExcelApp <| fun app -> 

            let workbook : Excel.Workbook = app.Workbooks.Open(inputFile)
            if fitWidth then 
                workbook.Sheets 
                    |> Seq.cast<Excel.Worksheet>
                    |> Seq.iter (fun (sheet:Excel.Worksheet) -> 
                        sheet.PageSetup.Zoom <- false
                        sheet.PageSetup.FitToPagesWide <- 1)
            else ()

            workbook.ExportAsFixedFormat( Type=Excel.XlFixedFormatType.xlTypePDF
                                        , Filename=outputFile
                                        , IncludeDocProperties=true
                                        , Quality = quality )
            workbook.Close (SaveChanges = false)
            Ok ()
        with
        | ex -> Error (sprintf "excelExportAsPdf failed '%s'" inputFile)



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
                         (inputFile:string) 
                         (outputFile:string) 
                         (searches:SearchList) : Result<unit,ErrMsg> = 
        try
            let workbook : Excel.Workbook = app.Workbooks.Open(inputFile)
            try 
                workbookFindReplace workbook searches
                workbook.SaveAs (Filename = outputFile)
                Ok ()
            with 
            | _ -> 
                workbook.Close (SaveChanges = false)
                Error (sprintf "excelFindReplace some error '%s'" inputFile) 
        with
        | _ -> Error (sprintf "excelFindReplace failed '%s'" inputFile)

