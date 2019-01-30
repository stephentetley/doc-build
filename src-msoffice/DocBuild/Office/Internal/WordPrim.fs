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
                        (quality:Word.WdExportOptimizeFor)
                        (inputFile:string) 
                        (outputFile:string) : Result<unit,ErrMsg> =
        try 
            let doc:(Word.Document) = app.Documents.Open(FileName = refobj inputFile)
            doc.ExportAsFixedFormat ( OutputFileName = outputFile
                                    , ExportFormat = Word.WdExportFormat.wdExportFormatPDF
                                    , OptimizeFor = quality)
            doc.Close (SaveChanges = refobj false)
            Ok ()
        with
        | _ -> Error (sprintf "wordExportAsPdf failed '%s'" inputFile)


    // ****************************************************************************
    // Find/Replace

    

    let private updateTableOfContents (doc:Word.Document)  : unit = 
        doc.TablesOfContents
            |> Seq.cast<Word.TableOfContents>
            |> Seq.iter (fun x -> x.Update ())

    /// TODO - I can't remember why this was so convoluted, it ought to be 
    /// simplified at some point.
    let private getHeadersOrFooters (proj:Word.Section -> Word.HeadersFooters) 
                                    (doc:Word.Document) : Word.HeaderFooter list = 
        Seq.foldBack (fun (section:Word.Section) (ac:Word.HeaderFooter list) ->
               let headers1 = proj section |> Seq.cast<Word.HeaderFooter>
               Seq.foldBack (fun x xs -> x::xs) headers1 ac)
               (doc.Sections |> Seq.cast<Word.Section>)
               []

    let private rangeFindReplace (search:string, replace:string) (range:Word.Range) : unit =
        range.Find.ClearFormatting ()
        range.Find.Execute (FindText = refobj search, 
                            ReplaceWith = refobj replace,
                            Replace = refobj Word.WdReplace.wdReplaceAll) |> ignore


    let private replacer (args :string * string) (doc:Word.Document) : unit = 
        let rngAll = doc.Range()
        rangeFindReplace args rngAll 
        let headers = getHeadersOrFooters (fun section -> section.Headers) doc
        let footers = getHeadersOrFooters (fun section -> section.Footers) doc
        List.iter (fun (header:Word.HeaderFooter) -> rangeFindReplace args header.Range)
                  (headers @ footers)
    


    // Note when debugging.
    //
    // The doc is traversed multiple times (once for each find-replace pair).
    // In practice this is not heinuous as the "traversal" is very shallow -
    // get the doc, then get its headers and footers.
    //
    // It can make debug output confusing though. 

    let documentFindReplace (searches:SearchList) 
                            (doc:Word.Document) : unit = 
        List.iter (fun search1 -> replacer search1 doc) searches 
        updateTableOfContents doc


    let private doubleQuote (s:string) : string = "\"" + s + "\""
    

    let wordFindReplace (app:Word.Application)
                        (searches:SearchList)
                        (inputFile:string) 
                        (outputFile:string) : Result<unit,ErrMsg> = 
        try
            let doc = app.Documents.Open(FileName = refobj inputFile)
            documentFindReplace searches doc 
            try 
                let outpath1 = doubleQuote outputFile
                doc.SaveAs (FileName = refobj outpath1)
                Ok ()
            with 
            | _ -> 
                doc.Close (SaveChanges = refobj false) |> ignore
                Error (sprintf "wordFindReplace some error '%s'" inputFile)
        with
        | _ -> Error (sprintf "wordFindReplace failed '%s'" inputFile)

