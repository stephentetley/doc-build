﻿[<AutoOpen>]
module DocMake.Tasks.DocPhotos


open System.IO
open System.Text.RegularExpressions
open Microsoft.Office.Interop

open DocMake.Base.Office


[<CLIMutable>]
type DocPhotosParams = 
    { 
        InputPaths : string list
        OutputFile : string
        ShowFileName : bool
    }

let DocPhotosDefaults = 
    { InputPaths = []
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


        
let getJPEGs (dirs:string list) : string list = 
    let re = new Regex("\.je?pg$", RegexOptions.IgnoreCase)
    let get1 = fun dir -> 
        Directory.GetFiles(dir) 
        |> Array.filter (fun s -> re.Match(s).Success)
        |> Array.toList
    List.collect get1 dirs
    

type StepFun = Word.Document -> string -> unit


let stepWithoutLabel : StepFun = appendPicture

let stepWithLabel : StepFun = 
    fun doc filename -> appendPicture doc filename
                        appendText doc <| Path.GetFileName filename

let processPhotos (doc:Word.Document) (action1:StepFun) (files:string list) : unit =
    let rec work zs = 
        match zs with 
        | [] -> ()
        | [x] -> action1 doc x
        | x :: xs -> action1 doc x
                     appendPageBreak doc
                     work xs
    work files
                            

let DocPhotos (setDocPhotosParams: DocPhotosParams -> DocPhotosParams) : unit =
    let opts = DocPhotosDefaults |> setDocPhotosParams
    let jpegs = getJPEGs opts.InputPaths
    let app = new Word.ApplicationClass (Visible = true)
    let stepFun = if opts.ShowFileName then stepWithLabel else stepWithoutLabel
    try 
        let doc = app.Documents.Add()
        processPhotos doc stepFun jpegs
        doc.SaveAs(FileName= refobj opts.OutputFile)
        doc.Close(SaveChanges = refobj false)
    finally 
        app.Quit ()
    




    



