// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

module DocMake.Base.SimpleDocOutput

open Microsoft.Office.Interop

open DocMake.Base.OfficeUtils

type Result<'a> = 
    | Failure of string
    | Success of 'a



// Word Output Monad
// Output is to a handle so this is not properly a writer monad
// (all output must be sequential)

type WordDoc = Word.Document

type DocOutput<'a> = DocOutput of (WordDoc -> 'a)

let inline private apply1 (ma : DocOutput<'a>) (handle:WordDoc) : 'a = 
    let (DocOutput f) = ma in f handle

let inline private unitM (x:'a) : DocOutput<'a> = DocOutput (fun _ -> x)


let inline private bindM (ma:DocOutput<'a>) (f : 'a -> DocOutput<'b>) : DocOutput<'b> =
    DocOutput (fun doc -> let a = apply1 ma doc in apply1 (f a) doc)




type DocOutputBuilder() = 
    member self.Return x = unitM x
    member self.Bind (p,f) = bindM p f
    member self.Zero () = unitM ()

let (docOutput:DocOutputBuilder) = new DocOutputBuilder()


// Common monadic operations
let fmapM (fn:'a -> 'b) (ma:DocOutput<'a>) : DocOutput<'b> = 
    DocOutput <| fun (doc:WordDoc) ->
        let ans = apply1 ma doc in fn ans

let mapM (fn:'a -> DocOutput<'b>) (xs:'a list) : DocOutput<'b list> = 
    let rec work ac ys = 
        match ys with
        | [] -> unitM <| List.rev ac
        | z :: zs -> bindM (fn z) (fun a -> work (a::ac) zs)
    work [] xs

let forM (xs:'a list) (fn:'a -> DocOutput<'b>) : DocOutput<'b list> = mapM fn xs

let mapMz (fn:'a -> DocOutput<'b>) (xs:'a list) : DocOutput<unit> = 
    let rec work ys = 
        match ys with
        | [] -> unitM ()
        | z :: zs -> bindM (fn z) (fun _ -> work zs)
    work xs

let forMz (xs:'a list) (fn:'a -> DocOutput<'b>) : DocOutput<unit> = mapMz fn xs

let traverseM (fn: 'a -> DocOutput<'b>) (source:seq<'a>) : DocOutput<seq<'b>> = 
    DocOutput <| fun doc ->
        Seq.map (fun x -> let mf = fn x in apply1 mf doc) source

// Need to be strict - hence use a fold
let traverseMz (fn: 'a -> DocOutput<'b>) (source:seq<'a>) : DocOutput<unit> = 
    DocOutput <| fun doc ->
        Seq.fold (fun ac x -> 
                    let ans  = apply1 (fn x) doc in ac) 
                 () 
                 source 


let runDocOutput (outputFile:string) (ma:DocOutput<'a>) : Result<'a> = 
    let app = new Word.ApplicationClass (Visible = true)
    try
        let doc = app.Documents.Add()
        let ans =  match ma with | DocOutput fn -> fn doc
        doc.SaveAs( FileName = refobj outputFile )
        doc.Close( SaveChanges = refobj false )
        app.Quit ()
        Success ans
    with
    | ex -> 
        app.Quit()
        Failure <| sprintf "runDocOutput failed, filename '%s'\nError message: %s" outputFile ex.Message

let runDocOutput2 (outputFile:string) (app:Word.Application) (ma:DocOutput<'a>) : Result<'a> = 
    try
        let doc = app.Documents.Add()
        let ans =  match ma with | DocOutput fn -> fn doc
        doc.SaveAs( FileName = refobj outputFile )
        doc.Close( SaveChanges = refobj false )
        Success ans
    with
    | ex -> 
        Failure <| sprintf "runDocOutput failed, filename '%s'\nError message: %s" outputFile ex.Message
  
  

let private getEndRange () : DocOutput<Word.Range> = 
    DocOutput <| fun doc ->
        doc.GoTo(What = refobj Word.WdGoToItem.wdGoToBookmark, Name = refobj "\EndOfDoc")

let private atEnd (fn:Word.Range -> 'a) : DocOutput<'a> = 
    docOutput { 
        let! rng = getEndRange ()
        let ans = fn rng
        return ans
        }

let tellPicture (fileName : string) : DocOutput<unit> =
    if System.IO.File.Exists(fileName) then 
        atEnd <| fun rng -> ignore <| rng.InlineShapes.AddPicture(FileName = fileName)
    else
        printfn "SimpleDocOutput - picture missing - %s" fileName
        unitM ()
       

let tellPageBreak () : DocOutput<unit> =
    atEnd <| fun rng -> rng.InsertBreak(Type = refobj Word.WdBreakType.wdPageBreak)

let tellText (text : string) : DocOutput<unit> =
    atEnd <| fun rng -> 
        rng.Text <- text

let tellStyledText (style : Word.WdBuiltinStyle) (text : string) : DocOutput<unit> = 
    atEnd <| fun rng -> 
        rng.Style <- refobj style
        rng.Text <- text

