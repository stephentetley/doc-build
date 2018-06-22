// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

[<RequireQualifiedAccess>]
module DocMake.Lib.DocPhotos


open System.IO
open System.Text.RegularExpressions
open Microsoft.Office.Interop

open DocMake.Base.FakeLike
open DocMake.Base.ImageMagickUtils
open DocMake.Base.SimpleDocOutput
open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis
open DocMake.Builder.WordBuilder



type DocPhotosOptions = 
    { CopyToSubDirectory: string
      DocTitle: string option
      ShowFileName: bool }


let private getJPEGs (dir:string) : string list = 
    let re = new Regex("\.je?pg$", RegexOptions.IgnoreCase)
    Directory.GetFiles(dir) 
        |> Array.filter (fun s -> re.Match(s).Success)
        |> Array.toList

    

let private copyJPEGs (jpgSrcDirectory:string) (outputSubDirectory:string) : BuildMonad<'res,string> = 
    localSubDirectory outputSubDirectory <| 
        buildMonad { 
            let jpegs = getJPEGs jpgSrcDirectory
            do! mapMz copyToWorkingDirectory jpegs
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

// TODO inputPaths should be paired with an optional rename procedure
let private photoDocImpl (getHandle:'res-> Word.Application) (opts:DocPhotosOptions) (inputPaths:string list) : BuildMonad<'res,WordDoc> =
    let docProc (jpegFolder:string) : DocOutput<unit>  = 
        let jpegs = getJPEGs jpegFolder
        let stepFun = if opts.ShowFileName then stepWithLabel else stepWithoutLabel
        docOutput { 
            do! addTitle opts.DocTitle
            do! insertPhotos stepFun jpegs
            }

    buildMonad { 
        do! mapMz (fun jpg -> copyJPEGs jpg opts.CopyToSubDirectory) inputPaths
        let! outDoc = freshDocument () |>> documentChangeExtension "pdf"
        let! app = asksU getHandle
        let! tempLoc = (fun d -> d </> opts.CopyToSubDirectory) <<| askWorkingDirectory ()
        let _ = runDocOutput2 outDoc.DocumentPath app (docProc tempLoc)
        return outDoc
        } 
    

type DocPhotos<'res> = 
    { docPhotos : DocPhotosOptions -> string list -> BuildMonad<'res, WordDoc> }

let makeAPI (getHandle:'res -> Word.Application) : DocPhotos<'res> = 
    { docPhotos = photoDocImpl getHandle }
