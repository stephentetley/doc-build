﻿module DocMake.Lib.DocFindReplace


open System.IO
// Open at .Interop rather than .Word then the Word API has to be qualified
open Microsoft.Office.Interop


open DocMake.Base.Common
open DocMake.Base.OfficeUtils
open DocMake.Builder.BuildMonad
open DocMake.Builder.Builders
open DocMake.Builder.Basis


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
// The doc is traversed multiple times (once for each find-replace pair).
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



// Ideally this would be a function from (something like) WordDoc -> WordBuild<WordDoc>
// Then we could compose / chain document transformers. 


// is this a public or private function?
let getTemplate (filePath:string) : WordBuild<WordDoc> =
    assertFile filePath |> fmapM (fun s -> {DocumentPath = s})


// What to do about outfile name?
// If we generate a tempfile, we can have a more compact pipeline
let docFindReplace (matches:SearchList)  (template:WordDoc) : WordBuild<WordDoc> =
    buildMonad { 
        let! outName = freshFileName ()
        let! app = askU ()
        process1 app template.DocumentPath outName matches
        return (makeDocument outName)
        }

    
