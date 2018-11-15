// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


module DocMake.PdfDoc

open DocMake.Base.Common
open DocMake.Base.RunProcess

/// Concat PDFs with Ghostscript
/// We favour Ghostscript because it lets us lower the print 
/// quality (and reduce the file size).

 
let private gsOptions (quality:PdfPrintQuality) : string =
    match ghostscriptPrintSetting quality with
    | "" -> @"-dBATCH -dNOPAUSE -q -sDEVICE=pdfwrite"
    | ss -> sprintf @"-dBATCH -dNOPAUSE -q -sDEVICE=pdfwrite -dPDFSETTINGS=%s" ss

let private gsOutputFile (fileName:string) : string = 
    sprintf "-sOutputFile=\"%s\"" fileName

let private gsInputFile (fileName:string) : string = sprintf "\"%s\"" fileName


/// Apparently we cannot send multiline commands to execProcess.
let private makeGsCommand (quality:PdfPrintQuality) (outputFile:string) (inputFiles: string list) : string = 
    let line1 = gsOptions quality + " " + gsOutputFile outputFile
    let rest = List.map gsInputFile inputFiles
    String.concat " " (line1 :: rest)



type PdfPath = string

type GhostscriptOptions = 
    { WorkingDirectory: string 
      GhostscriptExe: string 
      PrintQuality: PdfPrintQuality         // Make this Ghostscript specific?
    }

/// A PdfDoc is actually a list of Pdf files that are rendered 
/// to a single document with Ghostscript.
/// This means we have monodial concatenation.
type PdfDoc = 
    val Documents : PdfPath list

    new () = 
        { Documents = [] }

    new (filePath:PdfPath) = 
        { Documents = [filePath] }

    internal new (paths:PdfPath list ) = 
        { Documents = paths }


    member internal v.Body 
        with get() : PdfPath list = v.Documents

    member v.Save(options: GhostscriptOptions, outputPath: string) : unit = 
        let command = makeGsCommand options.PrintQuality outputPath v.Body
        match executeProcess options.WorkingDirectory options.GhostscriptExe command with
        | Choice2Of2 i when i = 0 -> ()
        | Choice2Of2 i -> 
            printfn "%s" command; failwithf "PdfDoc.Save - error code %i" i
        | Choice1Of2 msg -> 
            printfn "%s" command; failwithf "PdfDoc.Save - '%s'" msg



let emptyPdf : PdfDoc = new PdfDoc ()

let pdfDoc (path:PdfPath) : PdfDoc = new PdfDoc (filePath = path)

let (^^) (x:PdfDoc) (y:PdfDoc) : PdfDoc = 
    new PdfDoc(paths = x.Body @ y.Body)

let concat (docs:PdfDoc list) = 
    let xs = List.concat (List.map (fun (d:PdfDoc) -> d.Body) docs)
    new PdfDoc(paths = xs)