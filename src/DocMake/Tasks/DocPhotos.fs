// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

[<RequireQualifiedAccess>]
module DocMake.Tasks.DocPhotos


open System.IO
open System.Text.RegularExpressions
open Microsoft.Office.Interop

open DocMake.Base.FakeLike
open DocMake.Base.ImageMagickUtils
open DocMake.Base.SimpleDocOutput
open DocMake.Builder.BuildMonad
open DocMake.Builder.Document
open DocMake.Builder.Basis



type DocPhotosOptions = 
    { CopyToSubDirectory: string
      DocTitle: string option
      ShowFileName: bool }


let private getJPEGs1 (dirPath:string) : string list = 
    let re = new Regex("\.je?pg$", RegexOptions.IgnoreCase)
    if System.IO.DirectoryInfo(dirPath).Exists then 
        Directory.GetFiles(dirPath) 
            |> Array.filter (fun s -> re.Match(s).Success)
            |> Array.toList
    else []        

    

let private copyJPEGs (jpgSrcDirectory:string) (outputSubDirectory:string) : BuildMonad<'res,string> = 
    localSubDirectory outputSubDirectory <| 
        buildMonad { 
            let jpegs = getJPEGs1 jpgSrcDirectory
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
        let jpegs = getJPEGs1 jpegFolder
        let stepFun = if opts.ShowFileName then stepWithLabel else stepWithoutLabel
        docOutput { 
            do! addTitle opts.DocTitle
            do! insertPhotos stepFun jpegs
            }

    buildMonad { 
        do! mapMz (fun jpg -> copyJPEGs jpg opts.CopyToSubDirectory) inputPaths
        let! outDoc = freshDocument "docx"
        let! app = asksU getHandle
        let! jpegCopiesLoc = (fun d -> d </> opts.CopyToSubDirectory) <<| askWorkingDirectory ()
        match outDoc.GetPath with
        | None -> return zeroDocument
        | Some outPath -> 
            if System.IO.Directory.Exists(jpegCopiesLoc) && System.IO.Directory.GetFiles(jpegCopiesLoc).Length > 0 then
                let _ = runDocOutput2 outPath app (docProc jpegCopiesLoc)
                return outDoc
            else
                return zeroDocument
        } 
    

type DocPhotosApi<'res> = 
    { DocPhotos : DocPhotosOptions -> string list -> BuildMonad<'res, WordDoc> }

let makeAPI (getHandle:'res -> Word.Application) : DocPhotosApi<'res> = 
    { DocPhotos = photoDocImpl getHandle }
