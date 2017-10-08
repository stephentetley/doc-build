[<AutoOpen>]
module DocMake.Tasks.DocPhotos


open System
open Microsoft.Office.Interop

open DocMake.Utils.Office


[<CLIMutable>]
type DocPhotosParams = 
    { 
        InputPath : string
        OutputFile : string
        ShowFileName : bool
    }

let DocPhotosDefaults = 
    { InputPath = @""
      OutputFile = @"photos.docx"
      ShowFileName = true }


let getEndRange (doc:Word.Document) : Word.Range = 
    doc.GoTo(What = refobj Word.WdGoToItem.wdGoToBookmark, Name = refobj "\EndOfDoc")

let appendPicture (doc:Word.Document) (filename : string) : unit =
    let rng = getEndRange doc
    ignore <| rng.InlineShapes.AddPicture(FileName = filename)

let appendPageBreak (doc:Word.Document) : unit =
    let rng = getEndRange doc
    ignore <| rng.InsertBreak(Type = refobj Word.WdBreakType.wdPageBreak)

let appendText (doc:Word.Document) (text : string) : unit =
    let rng = getEndRange doc
    rng.Text <- text

let appendStyledText (doc:Word.Document) (sty : Word.WdBuiltinStyle) (text : string) = 
    let rng = getEndRange doc
    rng.Style <- refobj sty
    rng.Text <- text



// This is implemented as an object (because it is mutable) but actually 
// a Writer monad-like API might be better

type DocBuilder =
    val mutable private rnglast : Word.Range
    
    member private x.GotoEnd () = 
        match x.rnglast with
        | null -> ()
        | _ -> x.rnglast.InsertParagraphAfter ()
               let parent = x.rnglast.Document
               let pcount = parent.Paragraphs.Count 
               x.rnglast <- parent.Paragraphs.Item(pcount).Range
               

    member public x.Document  with get () : Word.Document = x.rnglast.Document

    member public x.AppendPicture (filename : string) = 
        match x.rnglast with
        | null -> ()
        | _ -> x.GotoEnd ()
               ignore <| x.rnglast.InlineShapes.AddPicture(FileName = filename)

    member public x.AppendPageBreak () = 
        match x.rnglast with
        | null -> ()
        | _ -> x.GotoEnd ()
               ignore <| x.rnglast.InsertBreak(Type = refobj Word.WdBreakType.wdPageBreak)   // wdPageBreak


    member public x.AppendTextParagraph (s : string) = 
        match x.rnglast with
        | null -> ()
        | _ -> x.GotoEnd ()
               ignore <| x.rnglast.Text <- s

    member public x.AppendStyledParagraph (sty : Word.WdBuiltinStyle) (s : string) = 
        match x.rnglast with
        | null -> ()
        | _ -> x.GotoEnd ()
               ignore <| x.rnglast.Style <- refobj sty
               ignore <| x.rnglast.Text <- s

    
    new (odoc : Word.Document) = 
        let ix = odoc.Paragraphs.Count
        let r1 : Word.Range = if ix > 0 then odoc.Paragraphs.Item(ix).Range
                              else odoc.Range(Start = ref (0 :> obj)) 
        { rnglast = r1 }
        

let DocPhotos (setDocPhotosParams: DocPhotosParams -> DocPhotosParams) : unit =
    let opts = DocPhotosDefaults |> setDocPhotosParams
    ()




    



