[<AutoOpen>]
module DocMake.Tasks.DocPhotos


open System.IO
open System.Text.RegularExpressions
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


        
let getJPEGs (dir:string)  : string [] = 
    let re = new Regex("\.je?pg$", RegexOptions.IgnoreCase)
    Directory.GetFiles(dir) |> Array.filter (fun s -> re.Match(s).Success)

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
    let jpegs = Array.toList <| getJPEGs opts.InputPath
    let app = new Word.ApplicationClass (Visible = true)
    let stepFun = if opts.ShowFileName then stepWithLabel else stepWithoutLabel
    try 
        let doc = app.Documents.Add()
        processPhotos doc stepFun jpegs
        doc.SaveAs(FileName= refobj opts.OutputFile)
        doc.Close(SaveChanges = refobj false)
    finally 
        app.Quit ()
    




    



