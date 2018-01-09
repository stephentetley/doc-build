[<AutoOpen>]
module DocMake.Tasks.DocFindReplace


open System.IO
// Open at .Interop rather than .Word then the Word API has to be qualified
open Microsoft.Office.Interop

open DocMake.Base.Common
open DocMake.Base.Json
open DocMake.Base.Office

// Ideally Tasks want error (warning) logging...

// If we encode the search list as JSON, we want a JSON parser that 
// does not interpret/type the dictionary.

type SearchList = List<string*string>



[<CLIMutable>]
type DocFindReplaceParams = 
    { 
        TemplateFile : string
        JsonMatchesFile : string
        OutputFile : string
    }


let DocFindReplaceDefaults = 
    { TemplateFile = @""
      JsonMatchesFile = @""
      OutputFile = @"findreplace.docx" }



let private replacer (x:Word.Document) (search:string) (replace:string) : bool = 
    let dstart = x.Content.Start
    let dend = x.Content.End
    let rangeall = x.Range(refobj dstart, refobj dend)
    rangeall.Find.ClearFormatting ()
    rangeall.Find.Execute (FindText = refobj search, 
                            ReplaceWith = refobj replace,
                            Replace = refobj Word.WdReplace.wdReplaceAll)

let private replaces (x:Word.Document) (zs:SearchList) : unit = 
    for z in zs do
        match z with | (a,b) -> ignore <| replacer x a b


let private process1 (app:Word.Application) (inpath:string) (outpath:string) (ss:SearchList) = 
    let doc = app.Documents.Open(FileName = refobj inpath)
    replaces doc ss
    // This should be wrapped in try...
    try 
        let outpath1 = doubleQuote outpath
        printfn "Outpath: %s" outpath1
        doc.SaveAs (FileName = refobj outpath1)
    finally 
        doc.Close (SaveChanges = refobj false)


let DocFindReplace (setDocFindReplaceParams: DocFindReplaceParams -> DocFindReplaceParams) : unit =
    let options = DocFindReplaceDefaults |> setDocFindReplaceParams
    if File.Exists(options.TemplateFile) && File.Exists(options.JsonMatchesFile)
    then
        let app = new Word.ApplicationClass (Visible = true)
        try 
            let matches = readJsonStringPairs options.JsonMatchesFile
            process1 app options.TemplateFile options.OutputFile matches
        finally 
            app.Quit ()
    else ()


