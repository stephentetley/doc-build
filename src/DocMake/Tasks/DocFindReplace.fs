[<AutoOpen>]
module DocMake.Tasks.DocFindReplace


open System.IO
open Microsoft.Office.Interop.Word

open DocMake.Utils.Common
open DocMake.Utils.Office

// Ideally Tasks want error (warning) logging...

// If we encode the search list as JSON, we want a JSON parser that 
// does not interpret/type the dictionary.

type SearchList = List<string*string>



[<CLIMutable>]
type DocFindReplaceParams = 
    { 
        InputFile : string
        OutputFile : string
        Searches : SearchList
    }


let DocFindReplaceDefaults = 
    { OutputFile = @"findreaplce.docx"
      InputFile = @""
      Searches = [] }



let replacer (x:Document) (search:string) (replace:string) : bool = 
    let dstart = x.Content.Start
    let dend = x.Content.End
    let rangeall = x.Range(refobj dstart, refobj dend)
    rangeall.Find.ClearFormatting ()
    rangeall.Find.Execute (FindText = refobj search, ReplaceWith = refobj replace)

let replaces (x:Document) (zs:SearchList) : unit = 
    for z in zs do
        match z with | (a,b) -> ignore <| replacer x a b


let process1 (app:Application) (inpath:string) (outpath:string) (ss:SearchList) = 
    let doc = app.Documents.Open(FileName = refobj inpath)
    replaces doc ss
    // This should be wrapped in try...
    try 
        let outpath1 = doubleQuote outpath
        printfn "Outpath: %s" outpath1
        doc.SaveAs (FileName = refobj outpath1)
    finally 
        doc.Close (SaveChanges = refobj false)

//type ModList = string * SearchList
//
//let processMany (inpath:string) (mods:ModList list) = 
//    let app = new ApplicationClass (Visible = false)
//    for m1 in mods do 
//        match m1 with | (outpath,ss) ->  process1 app inpath outpath ss
//    app.Quit ()

let DocFindReplace (setDocFindReplaceParams: DocFindReplaceParams -> DocFindReplaceParams) : unit =
    let parameters = DocFindReplaceDefaults |> setDocFindReplaceParams
    if File.Exists(parameters.InputFile) 
    then
        let app = new ApplicationClass (Visible = true)
        try 
            process1 app parameters.InputFile parameters.OutputFile parameters.Searches
        finally 
            app.Quit ()
    else ()



// TEMP - because intellisense is not working we put a test string in the module
// to check compilation is "successful".
let DocFindReplace_Teststring = "DOC_FIND_REPLACE"
