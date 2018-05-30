module DocMake.Lib.DocPhotos


open System.IO
open System.Text.RegularExpressions
open Microsoft.Office.Interop

open DocMake.Base.ImageMagickUtils
open DocMake.Base.SimpleDocOutput
open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis
open DocMake.Builder.Builders

// TODO - must optimize the size of the photos for Word.
// Which in turn means copy the photos.




let private getJPEGs (dirs:string list) : string list = 
    let re = new Regex("\.je?pg$", RegexOptions.IgnoreCase)
    let get1 = fun dir -> 
        Directory.GetFiles(dir) 
        |> Array.filter (fun s -> re.Match(s).Success)
        |> Array.toList
    List.collect get1 dirs
    

let private copyJPEGs (imgPaths:string list) : BuildMonad<'res,string> = 
    localSubDirectory "photos" <| 
        buildMonad { 
            do! mapMz copyToWorkingDirectory imgPaths
            let! cwd = askWorkingDirectory ()
            do optimizePhotos cwd 
            return cwd
        }

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
        

let private insertPhotos (action1:PictureFun) (files:string list) : DocOutput<unit> =
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
    let docProc (jpegFolder:string) : DocOutput<unit>  = 
        let jpegs = getJPEGs [jpegFolder]
        let stepFun = if showJpegFileName then stepWithLabel else stepWithoutLabel
        docOutput { 
            do! addTitle documentTitle
            do! insertPhotos stepFun jpegs
            }

    buildMonad { 
        let jpegInputs = getJPEGs inputPaths
        let! tempLoc = copyJPEGs jpegInputs
        let! outDoc = freshDocument ()
        let _ = runDocOutput outDoc.DocumentPath (docProc tempLoc)
        return outDoc
        }
    

