// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


[<RequireQualifiedAccess>]
module DocMake.Lib.PdfConcat


open DocMake.Base.Common
open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis
open DocMake.Builder.GhostscriptBuilder
 

// Concat PDFs with Ghostscript
 
let private makeGsOptions (quality:PdfPrintSetting) : string =
    match ghostscriptPrintSetting quality with
    | "" -> @"-dBATCH -dNOPAUSE -q -sDEVICE=pdfwrite"
    | ss -> sprintf @"-dBATCH -dNOPAUSE -q -sDEVICE=pdfwrite -dPDFSETTINGS=%s" ss



let private line1 (quality:PdfPrintSetting) (outputFile:string) : string =
    sprintf "%s -sOutputFile=\"%s\"" (makeGsOptions quality)  outputFile

let private lineQuote (name:string) : string = sprintf "  \"%s\"" name

//let private unlines (lines: string list) : string = String.concat "\n" lines
//let private unlinesC (lines: string list) : string = String.concat "^\n" lines
let private unlinesS (lines: string list) : string = String.concat " " lines


let private makeCmd (quality:PdfPrintSetting) (outputFile:string) (inputFiles: string list) : string = 
    let first = line1 quality outputFile
    let rest = List.map lineQuote inputFiles
    unlinesS <| first :: rest


let private pdfConcatImpl (getHandle:'res -> GsHandle) (inputFiles:PdfDoc list) : BuildMonad<'res,PdfDoc> = 
    let paths = List.map (fun (a:PdfDoc) -> a.DocumentPath) inputFiles
    buildMonad { 
        let! outDoc = freshDocument () |>> documentChangeExtension "pdf"
        let! quality = asksEnv (fun s -> s.PdfQuality)
        let! _ =  gsRunCommand getHandle <| makeCmd quality outDoc.DocumentPath paths
        return (castToPdfDoc outDoc)
    }


type PdfConcat<'res> = 
    { pdfConcat : PdfDoc list -> BuildMonad<'res, PdfDoc> }

let makeAPI (getHandle:'res-> GsHandle) : PdfConcat<'res> = 
    { pdfConcat = pdfConcatImpl getHandle }

    
  


