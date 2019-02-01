// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


/// Generally use Collection.* prefix with this module

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
        
        static member ( <<& ) (item:Document<'a>, col:Collection<'a>) = 
            match col with
            | Empty -> One item
            | _  -> Join (One(item), col)

        static member ( <<& ) (item:Document<'a> option, col:Collection<'a>) = 
            match item, col with
            | None, _ -> col
            | Some x, Empty -> One (x)
            | Some x, _ -> Join (One(x), col)

        static member ( &>> ) (col:Collection<'a>, item:Document<'a>) = 
            match col with
            | Empty -> One item
            | _  -> Join (col, One(item))

        static member ( &>> ) (col:Collection<'a>, items:Document<'a> list) = 
            let rec work xs acc =
                match xs with
                | [] -> acc
                | d :: ds -> work ds (acc &>> d)
            work items col 
                    
        static member ( &>> ) (col:Collection<'a>, item:Document<'a> option) = 
            match item, col with
            | None, _ -> col
            | Some x, Empty -> One (x)
            | Some x, _ -> Join (col, One(x))

        static member ( @@ ) (col1:Collection<'a>, col2:Collection<'a>) = 
            match col1, col2 with
            | Empty, y -> y
            | x, Empty -> x
            | x, y -> Join (x,y)


    type ViewL<'a> = 
        | EmptyL
        | ViewL of Document<'a> * Collection<'a>

    type ViewR<'a> = 
        | EmptyR
        | ViewR of Collection<'a> * Document<'a>


    /// Right-associative fold of a JoinList.
    /// In CPS form
    let private joinfoldr (op:Document<'a> -> 'ans -> 'ans) 
                          (initial:'ans) 
                          (source:Collection<'a>) : 'ans = 
        let rec work (src:Collection<'a>) (acc:'ans) (cont:'ans -> 'ans): 'ans = 
            match src with
            | Empty -> cont acc
            | One(a) -> cont (op a acc)
            | Join(t,u) -> 
                work u acc  (fun v1 -> 
                work t v1   (fun v2 -> 
                cont v2))
        work source initial (fun a -> a)

    /// Left-associative fold of a JoinList.
    /// In CPS form
    let private joinfoldl (op:Document<'a> -> 'ans -> 'ans) 
                          (initial:'ans) 
                          (source:Collection<'a>) : 'ans = 
        let rec work (src:Collection<'a>) (acc:'ans) (cont:'ans -> 'ans): 'ans = 
            match src with
            | Empty -> cont acc
            | One(a) -> cont (op a acc)
            | Join(t,u) -> 
                work t acc  (fun v1 -> 
                work u v1   (fun v2 -> 
                cont v2))
        work source initial (fun a -> a)      



    
    // ************************************************************************
    // Construction

    /// Create an empty Collection.
    let empty : Collection<'a> = Empty


    /// Create a singleton Collection.
    let singleton (item:Document<'a>) : Collection<'a> = One(item)

    /// Concat.
    let concat (col1:Collection<'a>) 
               (col2:Collection<'a>) : Collection<'a> =
        match col1, col2 with
        | Empty, y -> y
        | x, Empty -> x
        | x, y -> Join (x,y)

    /// Add a Document to the left.
    let cons (item:Document<'a>) (col:Collection<'a>) : Collection<'a> = 
        match col with
        | Empty -> One(item)
        | _ -> Join(One(item), col)

    /// Add a Document to the right.
    let snoc (col:Collection<'a>) (item:Document<'a>) : Collection<'a> = 
        match col with
        | Empty -> One(item)
        | _ -> Join(col, One(item))
        
    // ************************************************************************
    // Conversion

    /// Convert a Collection to a regular list.
    let toList (source:Collection<'a>) : Document<'a> list = 
        joinfoldr (fun x xs -> x :: xs) [] source

    /// Build a Collection from a regular list.
    let fromList (source:Document<'a> list) : Collection<'a> = 
        List.fold snoc Empty source

    // ************************************************************************
    // Views

    /// Access the left end of a Collection.
    ///
    /// Note this is not a cheap operation, the Collection must 
    /// be traversed down the left spine to find the leftmost node.
    let viewl (source:Collection<'a>) : ViewL<'a> = 
        let rec work (src:Collection<'a>) (cont:ViewL<'a> -> ViewL<'a>) : ViewL<'a>= 
            match src with
            | Empty -> cont EmptyL
            | One(a) -> cont (ViewL(a,Empty))
            | Join(t,u) -> 
                work t (fun v1 -> 
                match v1 with 
                | EmptyL -> work u id
                | ViewL(a,spineL) ->
                    cont (ViewL(a, concat spineL u)))
        work source id


    /// Access the right end of a Collection.
    ///
    /// Note this is not a cheap operation, the Collection must 
    /// be traversed down the right spine to find the rightmost node.
    let viewr (source:Collection<'a>) : ViewR<'a> = 
        let rec work (src:Collection<'a>) (cont:ViewR<'a> -> ViewR<'a>) : ViewR<'a>= 
            match src with
            | Empty -> cont EmptyR
            | One(a) -> cont (ViewR(Empty,a))
            | Join(t,u) -> 
                work u (fun v1 -> 
                match v1 with 
                | EmptyR -> work t id
                | ViewR(spineR, a) ->
                    cont (ViewR(concat t spineR, a)))
        work source id

    let map (fn:Document<'a> -> Document<'b>) 
             (collection:Collection<'a>) : Collection<'b> =
        List.map fn (toList collection) |> fromList

    let mapM (fn:Document<'a> -> DocMonad.DocMonad<'res, Document<'b>>) 
             (collection:Collection<'a>) : DocMonad.DocMonad<'res, Collection<'b>> =
        DocMonad.mapM fn (toList collection) |>> fromList
        


[<AutoOpen>]
module TypedCollection = 

    open DocBuild.Base.DocMonad

    type PdfCollection = Collection.Collection<PdfPhantom>


    /// All files must have .pdf extension
    let fromPdfList (pdfs:PdfFile list) : DocMonad<'res,PdfCollection> = 
        Collection.fromList pdfs |> dreturn


    type JpegCollection = Collection.Collection<JpegPhantom>

    /// All files must have .pdf extension
    let fromJpegList (jpegs:JpegFile list) : DocMonad<'res,JpegCollection> = 
        Collection.fromList jpegs |> dreturn