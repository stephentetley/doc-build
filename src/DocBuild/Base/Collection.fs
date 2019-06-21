// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


/// Generally use Collection.* prefix with this module

// Note
// We don't need collection to do very much, just build from 
// either end allowing cons and snoc of option<Document<'a>>,
// plus we need access to the elements when built.
// Currently this module is overkill.

namespace DocBuild.Base

[<AutoOpen>]
module Collection = 

    open DocBuild.Base


    [<Struct>]
    type Collection<'a> = 
        internal | Collection of documents : Document<'a> list

        member x.Documents 
            with get () : Document<'a> list = 
                let select (doc: Document<'a>) : Document<'a> option = 
                    Option.map (fun _ -> doc) doc.AbsolutePath
                match x with 
                | Collection(xs) -> xs |> List.choose select 
                    

        member x.DocumentTitles 
            with get () : string list = 
                let select (doc: Document<'a>) : string option = 
                    Option.map (fun _ -> doc.Title) doc.AbsolutePath
                match x with 
                | Collection(xs) -> xs |> List.choose select 
        

        member x.DocumentPaths
            with get () : string list = 
                match x with 
                | Collection(xs) -> 
                    xs |> List.choose (fun doc -> doc.AbsolutePath)

        static member ofList (documents : Document<'a> list) : Collection<'a> = 
            Collection documents

        static member concat (collections : Collection<'a> list) : Collection<'a> = 
            collections 
                |> List.map (fun x -> x.Documents)
                |> List.concat
                |> Collection


        static member singleton (doc : Document<'a>) : Collection<'a> = 
            Collection [doc]

    let ( ^^ ) (doc : Document<'a> ) (col : Collection<'a>) : Collection<'a> = 
        match col with
        | Collection(xs) -> Collection (doc :: xs)

    type PdfCollection = Collection<PdfPhantom>


    /// All files must have .pdf extension
    let fromPdfList (pdfs:PdfDoc list) : DocMonad<PdfCollection, 'userRes> = 
        Collection.ofList pdfs |> mreturn


    type JpegCollection = Collection<JpegPhantom>

    /// All files must have .pdf extension
    let fromJpegList (jpegs:JpegDoc list) : DocMonad<JpegCollection, 'userRes> = 
        Collection.ofList jpegs |> mreturn