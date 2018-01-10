[<AutoOpen>]
module DocMake.Tasks.DocFindReplace


open System.IO
// Open at .Interop rather than .Word then the Word API has to be qualified
open Microsoft.Office.Interop

open DocMake.Base.Common
open DocMake.Base.Json
open DocMake.Base.Office

// NOTE - Range.Text should be displayed with great caution
// (This pertains to DocMonad especially) 
// It will often contain "unprintable" that cause rendering "error" moving the
// cursor backwards etc.

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

let private getAllHFBySection (doc:Word.Document) (proj:Word.Section -> Word.HeadersFooters) : Word.HeaderFooter list = 
    Seq.foldBack (fun (section:Word.Section) (ac:Word.HeaderFooter list) ->
           let headers1 = proj section |> Seq.cast<Word.HeaderFooter>
           Seq.foldBack (fun x xs -> x::xs) headers1 ac)
           (doc.Sections |> Seq.cast<Word.Section>)
           []

let private replaceRange (range:Word.Range) (search:string) (replace:string) : unit =
    range.Find.ClearFormatting ()
    //let len = min 60 (range.Text.Length)
    //printfn "Range.Text: '%s'" range.Text
    //printfn "Range Chars (first %i): %A" len (List.take len << Seq.toList <| range.Text)
    ignore <| range.Find.Execute (FindText = refobj search, 
                            ReplaceWith = refobj replace,
                            Replace = refobj Word.WdReplace.wdReplaceAll)

let private replacer (doc:Word.Document) (search:string, replace:string) : unit =                      
    let rngAll = doc.Range()
    replaceRange rngAll search replace
    let headers = getAllHFBySection doc (fun section -> section.Headers)
    let footers = getAllHFBySection doc (fun section -> section.Footers)
    List.iter (fun (header:Word.HeaderFooter) -> replaceRange header.Range search replace)
              (headers @ footers)
    


// Note when debugging.
//
// The doc is traversed multiple times (once per find-replace pair).
// In practice this is not heinuous as the "traversal" is very shallow -
// get the doc, then get its headers and footers.
//
// It can make debug output confusing though. 

let private replaces (doc:Word.Document) (searches:SearchList) : unit = 
    List.iter (replacer doc) searches 


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


