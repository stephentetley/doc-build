module DocMake.Lib.XlsFindReplace


open System.IO
// Open at .Interop rather than .Excel then the Excel API has to be qualified
open Microsoft.Office.Interop


open DocMake.Base.Common
open DocMake.Base.OfficeUtils
open DocMake.Builder.BuildMonad
open DocMake.Builder.Builders
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



// What to do about outfile name?
// If we generate a tempfile, we can have a more compact pipeline
let xlsFindReplace (matches:SearchList)  (xlsDoc:ExcelDoc) : ExcelBuild<ExcelDoc> =
    buildMonad { 
        let! outName = freshFileName ()
        let! app = askU ()
        process1 xlsDoc.DocumentPath outName matches app
        return (makeDocument outName)
        }

    
