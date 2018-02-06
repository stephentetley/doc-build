module DocMake.Tasks.DocPhotos


open System.IO
open System.Text.RegularExpressions
open Microsoft.Office.Interop

open DocMake.Base.Office


[<CLIMutable>]
type DocPhotosParams = 
    { InputPaths: string list
      OutputFile: string
      ShowFileName: bool
      DocumentTitle: string option
    }

let DocPhotosDefaults = 
    { InputPaths = []
      OutputFile = @"photos.docx"
      ShowFileName = true
      DocumentTitle = None }


let private getEndRange (doc:Word.Document) : Word.Range = 
    doc.GoTo(What = refobj Word.WdGoToItem.wdGoToBookmark, Name = refobj "\EndOfDoc")

let private appendPicture (doc:Word.Document) (filename : string) : unit =
    let rng = getEndRange doc
    ignore <| rng.InlineShapes.AddPicture(FileName = filename)

let private appendPageBreak (doc:Word.Document) : unit =
    let rng = getEndRange doc
    ignore <| rng.InsertBreak(Type = refobj Word.WdBreakType.wdPageBreak)

let private appendText (doc:Word.Document) (text : string) : unit =
    let rng = getEndRange doc
    rng.Text <- text

let private appendStyledText (doc:Word.Document) (sty : Word.WdBuiltinStyle) (text : string) = 
    let rng = getEndRange doc
    rng.Style <- refobj sty
    rng.Text <- text


        
let private getJPEGs (dirs:string list) : string list = 
    let re = new Regex("\.je?pg$", RegexOptions.IgnoreCase)
    let get1 = fun dir -> 
        Directory.GetFiles(dir) 
        |> Array.filter (fun s -> re.Match(s).Success)
        |> Array.toList
    List.collect get1 dirs
    

type PictureFun = Word.Document -> string -> unit


let stepWithoutLabel : PictureFun = appendPicture

let stepWithLabel : PictureFun = 
    let makeCaption (fileName:string) : string = 
        sprintf "\n%s" (Path.GetFileName fileName)
    fun doc filename -> 
        appendPicture doc filename
        appendStyledText doc Word.WdBuiltinStyle.wdStyleNormal <| makeCaption filename

let private processPhotos (doc:Word.Document) (action1:PictureFun) (files:string list) : unit =
    let rec work zs = 
        match zs with 
        | [] -> ()
        | [x] -> action1 doc x
        | x :: xs -> action1 doc x
                     appendPageBreak doc
                     work xs
    work files


let private addTitle (doc:Word.Document) (optTitle:string option) : unit =
    match optTitle with
    | None -> ()
    | Some title -> 
        appendStyledText doc Word.WdBuiltinStyle.wdStyleTitle (title + "\n\n")


let DocPhotos (setDocPhotosParams: DocPhotosParams -> DocPhotosParams) : unit =
    let opts = DocPhotosDefaults |> setDocPhotosParams
    let jpegs = getJPEGs opts.InputPaths
    let app = new Word.ApplicationClass (Visible = true)
    let stepFun = if opts.ShowFileName then stepWithLabel else stepWithoutLabel
    try 
        let doc = app.Documents.Add()
        addTitle doc opts.DocumentTitle
        processPhotos doc stepFun jpegs
        // File must not exist...
        doc.SaveAs(FileName= refobj opts.OutputFile)
        doc.Close(SaveChanges = refobj false)
    finally 
        app.Quit ()
    




    



