module DocMake.Tasks.DocPhotos


open System.IO
open System.Text.RegularExpressions
open Microsoft.Office.Interop

open DocMake.Base.OfficeUtils
open DocMake.Base.SimpleDocOutput

[<CLIMutable>]
type DocPhotosParams = 
    { InputPaths: string list
      OutputFile: string
      ShowFileName: bool
      DocumentTitle: string option }

let DocPhotosDefaults = 
    { InputPaths = []
      OutputFile = @"photos.docx"
      ShowFileName = true
      DocumentTitle = None }


        
let private getJPEGs (dirs:string list) : string list = 
    let re = new Regex("\.je?pg$", RegexOptions.IgnoreCase)
    let get1 = fun dir -> 
        Directory.GetFiles(dir) 
        |> Array.filter (fun s -> re.Match(s).Success)
        |> Array.toList
    List.collect get1 dirs
    

type PictureFun = string -> DocOutput<unit>


let stepWithoutLabel : PictureFun = 
    fun jpegPath -> tellPicture jpegPath

let stepWithLabel : PictureFun = 
    fun jpegPath -> 
        let caption = sprintf "\n%s" <| Path.GetFileName jpegPath
        docOutput {  
            do! tellPicture jpegPath
            do! tellStyledText Word.WdBuiltinStyle.wdStyleNormal caption 
            }

let private addTitle  (optTitle:string option) : DocOutput<unit> =
    match optTitle with
    | None -> docOutput { return () }
    | Some title -> 
        docOutput { 
            do! tellStyledText Word.WdBuiltinStyle.wdStyleTitle (title + "\n\n")
            }

let private processPhotos (action1:PictureFun) (files:string list) : DocOutput<unit> =
    // Don't add page break to last, hence use direct recursion
    
    let rec work zs = 
        match zs with 
        | [] -> docOutput { return () }
        | [x] -> action1 x
        | x :: xs -> 
            docOutput { 
                do! action1 x
                do! tellPageBreak ()
                do! work xs }
    work files




let DocPhotos (setDocPhotosParams: DocPhotosParams -> DocPhotosParams) : unit =
    let opts = DocPhotosDefaults |> setDocPhotosParams
    let jpegs = getJPEGs opts.InputPaths
    let stepFun = if opts.ShowFileName then stepWithLabel else stepWithoutLabel
    let procM = 
        docOutput { 
            do! addTitle opts.DocumentTitle
            do! processPhotos stepFun jpegs
            }
    try 
        runDocOutput opts.OutputFile procM
    finally 
        ()
    




    



