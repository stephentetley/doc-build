// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

[<RequireQualifiedAccess>]
module DocMake.Tasks.XlsFindReplace


open System.IO
// Open at .Interop rather than .Excel then the Excel API has to be qualified
open Microsoft.Office.Interop


open DocMake.Base.Common
open DocMake.Builder.BuildMonad
open DocMake.Builder.Document
open DocMake.Builder.Basis


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


let private process1 (inpath:string) (outpath:string) (ss:SearchList) (app:Excel.Application) = 
    // This should be wrapped in try...
    let workbook : Excel.Workbook = app.Workbooks.Open(inpath)
    try 
        replaces workbook ss
        printfn "Outpath: %s" outpath
        workbook.SaveAs (Filename = outpath)
    finally 
        workbook.Close (SaveChanges = false)

/// TODO - this should assert the file extension is *.xlsx or *.xlsm...
let private getTemplateImpl (filePath:string) : BuildMonad<'res,ExcelDoc> =
    assertFile filePath |>> (fun s -> {DocumentPath = s})

// What to do about outfile name?
// If we generate a tempfile, we can have a more compact pipeline
let private xlsFindReplaceImpl (getHandle:'res-> Excel.Application) (matches:SearchList)  (xlsDoc:ExcelDoc) : BuildMonad<'res,ExcelDoc> =
    withXlsxNamer <| buildMonad { 
        let! outName = freshFileName ()
        let! app = asksU getHandle
        process1 xlsDoc.DocumentPath outName matches app
        return (makeDocument outName |> castToXls)
        }

    
    
type XlsFindReplace<'res> = 
    { XlsFindReplace : SearchList -> ExcelDoc -> BuildMonad<'res,ExcelDoc>
      GetTemplateXls: string -> BuildMonad<'res,ExcelDoc> }

let makeAPI (getHandle:'res-> Excel.Application) : XlsFindReplace<'res> = 
    { XlsFindReplace = xlsFindReplaceImpl getHandle 
      GetTemplateXls = getTemplateImpl }