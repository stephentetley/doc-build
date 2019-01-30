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
                         (fitWidth:bool)
                         (quality:Excel.XlFixedFormatQuality)
                         (inputFile:string)
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

    let private sheetFindReplace (search:string, replace:string) 
                                 (worksheet:Excel.Worksheet) : unit = 
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

    
    let private sheetFindReplaceList (searches:SearchList) (worksheet:Excel.Worksheet) : unit =                      
        List.iter (fun search1 -> sheetFindReplace search1 worksheet) searches


    let workbookFindReplace (searches:SearchList) 
                            (workbook:Excel.Workbook)  : unit = 
        workbook.Worksheets 
            |> Seq.cast<Excel.Worksheet> 
            |> Seq.iter (fun (sheet:Excel.Worksheet) -> sheetFindReplaceList searches sheet )


    let excelFindReplace (app:Excel.Application) 
                         (searches:SearchList)
                         (inputFile:string) 
                         (outputFile:string) : Result<unit,ErrMsg> = 
        try
            let workbook : Excel.Workbook = app.Workbooks.Open(inputFile)
            try 
                workbookFindReplace searches workbook 
                workbook.SaveAs (Filename = outputFile)
                Ok ()
            with 
            | _ -> 
                workbook.Close (SaveChanges = false)
                Error (sprintf "excelFindReplace some error '%s'" inputFile) 
        with
        | _ -> Error (sprintf "excelFindReplace failed '%s'" inputFile)

