// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

[<RequireQualifiedAccess>]
module DocMake.Tasks.DocFindReplace


open System.IO
// Open at .Interop rather than .Word then the Word API has to be qualified
open Microsoft.Office.Interop


open DocMake.Base.Common
open DocMake.Base.OfficeUtils
open DocMake.Builder.BuildMonad
open DocMake.Builder.Document
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


let private updateToCs (doc:Word.Document)  : unit = 
    doc.TablesOfContents
        |> Seq.cast<Word.TableOfContents>
        |> Seq.iter (fun x -> x.Update ())


let private process1  (inpath:string) (outpath:string) (ss:SearchList) (app:Word.Application) = 
    let doc = app.Documents.Open(FileName = refobj inpath)
    replaces doc ss
    updateToCs doc
    // This should be wrapped in try...
    try 
        let outpath1 = doubleQuote outpath
        printfn "Outpath: %s" outpath1
        doc.SaveAs (FileName = refobj outpath1)
    finally 
        doc.Close (SaveChanges = refobj false)




/// TODO - this should assert the file extension is *.doc or *.docx...
let private getTemplateImpl (filePath:string) : BuildMonad<'res,WordDoc> =
    assertFile filePath |>> makeDocument


// What to do about outfile name?
// If we generate a tempfile, we can have a more compact pipeline
let private docFindReplaceImpl (getHandle:'res-> Word.Application) (matches:SearchList)  (template:WordDoc) : BuildMonad<'res,WordDoc> =
    buildMonad { 
        let! outName = freshFileName "docx"
        let! app = asksU getHandle
        match template.GetPath with
        | None -> return zeroDocument
        | Some templatePath -> 
            process1 templatePath outName matches app
            return (makeDocument outName)
        }


    
type DocFindReplaceApi<'res> = 
    { DocFindReplace : SearchList -> WordDoc -> BuildMonad<'res,WordDoc>
      GetTemplateDoc: string -> BuildMonad<'res,WordDoc> }

let makeAPI (getHandle:'res-> Word.Application) : DocFindReplaceApi<'res> = 
    { DocFindReplace = docFindReplaceImpl getHandle
      GetTemplateDoc = getTemplateImpl }


