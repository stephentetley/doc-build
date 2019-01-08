// Copyright (c) Stephen Tetley 2018, 2019
// License: BSD 3 Clause

namespace DocBuild.Office.MsoWord



[<AutoOpen>]
module MsoWord = 

    open Microsoft.Office.Interop
    open DocBuild.Base.Common


    let internal rbox (x:'a) : ref<obj> = ref (x :> obj)


    let internal withWordApp (operation:Word.Application -> 'a) : 'a = 
        let app = new Word.ApplicationClass (Visible = true) :> Word.Application
        let result = operation app
        app.Quit()
        result


    let updateTableOfContents (doc:Word.Document)  : unit = 
        doc.TablesOfContents
            |> Seq.cast<Word.TableOfContents>
            |> Seq.iter (fun x -> x.Update ())


    // ****************************************************************************
    // Find/Replace


    /// TODO - I can't remember why this was so convoluted, it ought to be 
    /// simplified at some point.
    let private getHeadersOrFooters (doc:Word.Document) 
                                    (proj:Word.Section -> Word.HeadersFooters) : Word.HeaderFooter list = 
        Seq.foldBack (fun (section:Word.Section) (ac:Word.HeaderFooter list) ->
               let headers1 = proj section |> Seq.cast<Word.HeaderFooter>
               Seq.foldBack (fun x xs -> x::xs) headers1 ac)
               (doc.Sections |> Seq.cast<Word.Section>)
               []

    let private rangeFindReplace (range:Word.Range) (search:string) (replace:string) : unit =
        range.Find.ClearFormatting ()
        range.Find.Execute (FindText = rbox search, 
                            ReplaceWith = rbox replace,
                            Replace = rbox Word.WdReplace.wdReplaceAll) |> ignore


    let private replacer (doc:Word.Document) (search:string, replace:string) : unit =                      
        let rngAll = doc.Range()
        rangeFindReplace rngAll search replace
        let headers = getHeadersOrFooters doc (fun section -> section.Headers)
        let footers = getHeadersOrFooters doc (fun section -> section.Footers)
        List.iter (fun (header:Word.HeaderFooter) -> rangeFindReplace header.Range search replace)
                  (headers @ footers)
    


    // Note when debugging.
    //
    // The doc is traversed multiple times (once for each find-replace pair).
    // In practice this is not heinuous as the "traversal" is very shallow -
    // get the doc, then get its headers and footers.
    //
    // It can make debug output confusing though. 

    let documentFindReplace (doc:Word.Document) 
                            (searches:SearchList) : unit = 
        List.iter (replacer doc) searches 
        updateTableOfContents doc


    let private doubleQuote (s:string) : string = "\"" + s + "\""

    let wordFindReplace (app:Word.Application) 
                        (inpath:string) 
                        (outpath:option<string>) 
                        (searches:SearchList) : unit = 
        let doc = app.Documents.Open(FileName = rbox inpath)
        documentFindReplace doc searches
        try 
            match outpath with 
            | None -> doc.Save()
            | Some filename -> 
                let outpath1 = doubleQuote filename
                printfn "Outpath: %s" outpath1
                doc.SaveAs (FileName = rbox outpath1)
        finally 
            doc.Close (SaveChanges = rbox false)

