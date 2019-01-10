// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause

namespace DocBuild.Office.Internal


[<AutoOpen>]
module WordPrim = 

    open Microsoft.Office.Interop
    open DocBuild.Base
    open DocBuild.Office.Internal

    let internal withWordApp (operation:Word.Application -> 'a) : 'a = 
        let app = new Word.ApplicationClass (Visible = true) :> Word.Application
        let result = operation app
        app.Quit()
        result




    // ****************************************************************************
    // Export to Pdf

    let wordExportAsPdf (app:Word.Application) 
                    (inputFile:string) 
                    (quality:Word.WdExportOptimizeFor)
                    (outputFile:string) : Result<unit,ErrMsg> =
        try 
            let doc:(Word.Document) = app.Documents.Open(FileName = refobj inputFile)
            doc.ExportAsFixedFormat ( OutputFileName = outputFile
                                    , ExportFormat = Word.WdExportFormat.wdExportFormatPDF
                                    , OptimizeFor = quality)
            doc.Close (SaveChanges = refobj false)
            Ok ()
        with
        | ex -> Error (sprintf "exportAsPdf failed '%s'" inputFile)


    // ****************************************************************************
    // Find/Replace

    

    let private updateTableOfContents (doc:Word.Document)  : unit = 
        doc.TablesOfContents
            |> Seq.cast<Word.TableOfContents>
            |> Seq.iter (fun x -> x.Update ())

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
        range.Find.Execute (FindText = refobj search, 
                            ReplaceWith = refobj replace,
                            Replace = refobj Word.WdReplace.wdReplaceAll) |> ignore


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
                        (inputFile:string) 
                        (outputFile:string) 
                        (searches:SearchList) : unit = 
        let doc = app.Documents.Open(FileName = refobj inputFile)
        documentFindReplace doc searches
        try 
            let outpath1 = doubleQuote outputFile
            doc.SaveAs (FileName = refobj outpath1)
        finally 
            doc.Close (SaveChanges = refobj false)

