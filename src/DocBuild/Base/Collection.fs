// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


/// Generally use Collection.* prefix with this module

// Note
// We don't need collection to do very much, just build from 
// either end allowing cons and snoc of option<Document<'a>>,
// plus we need access to the elements when built.
// Currently this module is overkill.

namespace DocBuild.Base


module Collection = 

    open DocBuild.Base
    open DocBuild.Base.DocMonadOperators


    /// A Collection is a so called JoinList (unbalanced binary tree)
    /// of documents. It allows efficient addition to the right (snocing)
    /// and efficient append.
    type Collection<'a> = 
        private 
            | Empty
            | One of Document<'a>
            | Join of Collection<'a> * Collection<'a>
        
        
    /// Left-associative fold of a JoinList.
    /// In CPS form
        member v.fold (op:Document<'a> -> 'ans -> 'ans) 
                      (initial:'ans)  : 'ans = 
            let rec work (src:Collection<'a>) 
                         (acc:'ans) 
                         (cont:'ans -> 'ans): 'ans = 
                match src with
                | Empty -> cont acc
                | One(a) -> cont (op a acc)
                | Join(t,u) -> 
                    work t acc (fun v1 -> 
                    work u v1 cont)
            work v initial (fun a -> a)   

            /// Right-associative fold of a JoinList.
            /// In CPS form
        member v.foldBack (op:Document<'a> -> 'ans -> 'ans) 
                          (initial:'ans)  : 'ans = 
            let rec work (src:Collection<'a>) (acc:'ans) (cont:'ans -> 'ans): 'ans = 
                match src with
                | Empty -> cont acc
                | One(a) -> cont (op a acc)
                | Join(t,u) -> 
                    work u acc (fun v1 -> 
                    work t v1 cont)
            work v initial (fun a -> a)

        member v.Elements
            with get () : Document<'a> list = v.foldBack (fun a ac -> a::ac) []

        static member ( ^^& ) (item:Document<'a>, col:Collection<'a>) = 
            match col with
            | Empty -> One item
            | _  -> Join (One(item), col)

        static member ( ^^& ) (item:Document<'a> option, col:Collection<'a>) = 
            match item, col with
            | None, _ -> col
            | Some x, Empty -> One (x)
            | Some x, _ -> Join (One(x), col)

        static member ( &^^ ) (col:Collection<'a>, item:Document<'a>) = 
            match col with
            | Empty -> One item
            | _  -> Join (col, One(item))

        static member ( &^^ ) (col:Collection<'a>, items:Document<'a> list) = 
            let rec work xs acc =
                match xs with
                | [] -> acc
                | d :: ds -> work ds (acc &^^ d)
            work items col 
                    
        static member ( &^^ ) (col:Collection<'a>, item:Document<'a> option) = 
            match item, col with
            | None, _ -> col
            | Some x, Empty -> One (x)
            | Some x, _ -> Join (col, One(x))

        static member ( @@ ) (col1:Collection<'a>, col2:Collection<'a>) = 
            match col1, col2 with
            | Empty, y -> y
            | x, Empty -> x
            | x, y -> Join (x,y)

 


    
    // ************************************************************************
    // Construction

    /// Create an empty Collection.
    let empty : Collection<'a> = Empty


    /// Create a singleton Collection.
    let singleton (item:Document<'a>) : Collection<'a> = One(item)

    /// Build a Collection from a regular list.
    let fromList (source:Document<'a> list) : Collection<'a> = 
        List.fold (&^^) Empty source


    /// Concat.
    let concat (col1:Collection<'a>) 
               (col2:Collection<'a>) : Collection<'a> =
        match col1, col2 with
        | Empty, y -> y
        | x, Empty -> x
        | x, y -> Join (x,y)



[<AutoOpen>]
module TypedCollection = 

    open DocBuild.Base.DocMonad

    type PdfCollection = Collection.Collection<PdfPhantom>


    /// All files must have .pdf extension
    let fromPdfList (pdfs:PdfDoc list) : DocMonad<'res,PdfCollection> = 
        Collection.fromList pdfs |> mreturn


    type JpegCollection = Collection.Collection<JpegPhantom>

    /// All files must have .pdf extension
    let fromJpegList (jpegs:JpegDoc list) : DocMonad<'res,JpegCollection> = 
        Collection.fromList jpegs |> mreturn