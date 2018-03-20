module DocMake.Tasks.XlsFindReplace

open System.IO
open Microsoft.Office.Interop

open Fake
open Fake.Core

open DocMake.Base.Common
open DocMake.Base.Json
open DocMake.Base.Office

// NOTE - Range.Text should be displayed with great caution
// (This pertains to DocMonad especially) 
// It will often contain "unprintable" that cause rendering "error" moving the
// cursor backwards etc.

type SearchList = List<string*string>

// TODO - should probably change the representation of matches to SearchList (rather than JSON).
// The client can then work out weather or not it is serialized.

[<CLIMutable>]
type XlsFindReplaceParams = 
    { TemplateFile: string
      JsonMatchesFile: string
      OutputFile: string }


let XlsFindReplaceDefaults = 
    { TemplateFile = @""
      JsonMatchesFile = @""
      OutputFile = @"findreplace.xlsx" }


//let private replaceRange (range:Word.Range) (search:string) (replace:string) : unit =
//    range.Find.ClearFormatting ()
//    ignore <| range.Find.Execute (FindText = refobj search, 
//                                    ReplaceWith = refobj replace,
//                                    Replace = refobj Word.WdReplace.wdReplaceAll)

let private sheetReplacer1 (worksheet:Excel.Worksheet) (search:string, replace:string) : unit =                      
    ()
    //replaceRange rngAll search replace
    //List.iter (fun (header:Word.HeaderFooter) -> replaceRange header.Range search replace)
    //          (headers @ footers)
    


// Note when debugging.
//
// The doc is traversed multiple times (once per find-replace pair).
// In practice this is not heinuous as the "traversal" is very shallow -
// get the doc, then get its headers and footers.
//
// It can make debug output confusing though. 

let private replaces (workbook:Excel.Workbook) (searches:SearchList) : unit = 
    workbook.Worksheets 
        |> Seq.cast<Excel.Worksheet> 
        |> Seq.iter (fun (sheet:Excel.Worksheet) -> printfn "Worksheet: %s" sheet.Name)


let private process1 (app:Excel.Application) (inpath:string) (outpath:string) (ss:SearchList) = 
    // This should be wrapped in try...
    let workbook : Excel.Workbook = app.Workbooks.Open(inpath)
    try 
        replaces workbook ss
        let outpath1 = doubleQuote outpath
        printfn "Outpath: %s" outpath1
        workbook.SaveAs (Filename = refobj outpath1)
    finally 
        workbook.Close (SaveChanges = refobj false)


let DocFindReplace (setXlsFindReplaceParams: XlsFindReplaceParams -> XlsFindReplaceParams) : unit =
    let options = XlsFindReplaceDefaults |> setXlsFindReplaceParams
    match File.Exists(options.TemplateFile), File.Exists(options.JsonMatchesFile) with
    | true, true ->
        let app = new Excel.ApplicationClass(Visible = true)
        try 
            let matches = readJsonStringPairs options.JsonMatchesFile
            process1 app options.TemplateFile options.OutputFile matches
        finally 
            app.Quit ()
    | false, _ ->  
        Trace.traceError <| sprintf "XlsFindReplace --- missing template file '%s'" options.TemplateFile
        failwith "XlsFindReplace --- missing template file"
    | _, _ -> 
        Trace.traceError <| sprintf "XlsFindReplace --- missing matches file '%s'" options.JsonMatchesFile
        failwith "XlsFindReplace --- missing matches file"


