// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


[<RequireQualifiedAccess>]
module DocMake.Tasks.PdfConcat


open DocMake.Base.Common
open DocMake.Builder.BuildMonad
open DocMake.Builder.Document
open DocMake.Builder.ShellHooks
open DocMake.Builder.GhostscriptRunner

/// Concat PDFs with Ghostscript
/// We favour Ghostscript because it lets us lower the print 
/// quality (and reduce the file size).

 
let private makeGsOptions (quality:PdfPrintQuality) : string =
    match ghostscriptPrintSetting quality with
    | "" -> @"-dBATCH -dNOPAUSE -q -sDEVICE=pdfwrite"
    | ss -> sprintf @"-dBATCH -dNOPAUSE -q -sDEVICE=pdfwrite -dPDFSETTINGS=%s" ss



let private line1 (quality:PdfPrintQuality) (outputFile:string) : string =
    sprintf "%s -sOutputFile=\"%s\"" (makeGsOptions quality)  outputFile

let private lineQuote (name:string) : string = sprintf "  \"%s\"" name

//let private unlines (lines: string list) : string = String.concat "\n" lines
//let private unlinesC (lines: string list) : string = String.concat "^\n" lines
let private unlinesS (lines: string list) : string = String.concat " " lines


let private makeCmd (quality:PdfPrintQuality) (outputFile:string) (inputFiles: string list) : string = 
    let first = line1 quality outputFile
    let rest = List.map lineQuote inputFiles
    unlinesS <| first :: rest


let private pdfConcatImpl (getHandle:'res -> GsHandle) 
            (inputFiles:PdfDoc list) : BuildMonad<'res,PdfDoc> = 
    let paths = 
        List.choose id <| List.map (fun (a:PdfDoc) -> a.GetPath) inputFiles

    match paths with
    | [] -> breturn zeroDocument
    | _ -> 
        buildMonad { 
            let! outDoc = freshDocument "pdf"
            let! quality = asksEnv (fun s -> s.PdfQuality)
            match outDoc.GetPath with
            | None -> return zeroDocument
            | Some outPath -> 
                let! _ = gsRunCommand getHandle (makeCmd quality outPath paths)
                return outDoc
        }
  


type PdfConcatApi<'res> = 
    { PdfConcat : PdfDoc list -> BuildMonad<'res, PdfDoc> }

let makeAPI (getHandle:'res-> GsHandle) : PdfConcatApi<'res> = 
    { PdfConcat = pdfConcatImpl getHandle }


// ****************************************************************************

/// New API

let pdfConcat (inputFiles:PdfDoc list) (outfileName:string) : GhostscriptRunner<PdfDoc> = 
    let paths = 
        List.choose id <| List.map (fun (a:PdfDoc) -> a.GetPath) inputFiles

    match paths with
    | [] -> mreturn zeroDocument
    | _ -> 
        ghostscriptRunner { 
            let! outDoc = liftBM <| freshDocument "pdf"
            let! quality = liftBM <| asksEnv (fun s -> s.PdfQuality)
            match outDoc.GetPath with
            | None -> return zeroDocument
            | Some outPath -> 
                let! _ = ghostscriptExec (makeCmd quality outPath paths)
                return outDoc
        }
    
  


