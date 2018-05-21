module DocMake.Lib.DocPhotos


open System.IO
open System.Text.RegularExpressions
open Microsoft.Office.Interop


open DocMake.Base.SimpleDocOutput
open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis
open DocMake.Builder.Builders

        
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
    | Some title -> tellStyledText Word.WdBuiltinStyle.wdStyleTitle (title + "\n\n")
        

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


let photoDoc (documentTitle: string option) (showJpegFileName:bool) (inputPaths:string list) : WordBuild<WordDoc> =
    let docProc:DocOutput<unit>  = 
        let stepFun = if showJpegFileName then stepWithLabel else stepWithoutLabel
        docOutput { 
            do! addTitle documentTitle
            do! processPhotos stepFun inputPaths
            }
    buildMonad { 
        let! outDoc = freshDocument ()
        let _ = runDocOutput outDoc.DocumentPath docProc
        return outDoc
        }
    




    



