module DocMake.Tasks.DocFindReplace


open System.IO
// Open at .Interop rather than .Word then the Word API has to be qualified
open Microsoft.Office.Interop

open Fake
open Fake.Core

open DocMake.Base.Common
open DocMake.Base.JsonUtils
open DocMake.Base.OfficeUtils

// NOTE - Range.Text should be displayed with great caution
// (This pertains to DocMonad especially) 
// It will often contain "unprintable" that cause rendering "error" moving the
// cursor backwards etc.





[<CLIMutable>]
type DocFindReplaceParams = 
    { TemplateFile: string
      Matches: SearchList
      OutputFile: string }


let DocFindReplaceDefaults = 
    { TemplateFile = @""
      Matches = []
      OutputFile = @"findreplace.docx" }

let private getHeadersOrFooters (doc:Word.Document) (proj:Word.Section -> Word.HeadersFooters) : Word.HeaderFooter list = 
    Seq.foldBack (fun (section:Word.Section) (ac:Word.HeaderFooter list) ->
           let headers1 = proj section |> Seq.cast<Word.HeaderFooter>
           Seq.foldBack (fun x xs -> x::xs) headers1 ac)
           (doc.Sections |> Seq.cast<Word.Section>)
           []

let private replaceRange (range:Word.Range) (search:string) (replace:string) : unit =
    range.Find.ClearFormatting ()
    ignore <| range.Find.Execute (FindText = refobj search, 
                                    ReplaceWith = refobj replace,
                                    Replace = refobj Word.WdReplace.wdReplaceAll)

let private replacer (doc:Word.Document) (search:string, replace:string) : unit =                      
    let rngAll = doc.Range()
    replaceRange rngAll search replace
    let headers = getHeadersOrFooters doc (fun section -> section.Headers)
    let footers = getHeadersOrFooters doc (fun section -> section.Footers)
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
    match File.Exists(options.TemplateFile) with
    | true ->
        let app = new Word.ApplicationClass (Visible = true)
        try 
            process1 app options.TemplateFile options.OutputFile options.Matches
        finally 
            app.Quit ()
    | false ->  
        Trace.traceError <| sprintf "DocFindReplace --- missing template file '%s'" options.TemplateFile
        failwith "DocFindReplace --- missing template file"

/// Version of DocFindReplace with a passed in reference to Word
let BatchDocFindReplace (app:Word.Application) (setDocFindReplaceParams: DocFindReplaceParams -> DocFindReplaceParams) : unit =
    let options = DocFindReplaceDefaults |> setDocFindReplaceParams
    match File.Exists(options.TemplateFile) with
    | true ->
        try 
            process1 app options.TemplateFile options.OutputFile options.Matches
        with
        | ex -> 
            Trace.traceError <| sprintf "BatchDocFindReplace --- exception '%s'" ex.Message
    | false ->  
        Trace.traceError <| sprintf "BatchDocFindReplace --- missing template file '%s'" options.TemplateFile
        failwith "BatchDocFindReplace --- missing template file"

