// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


module DocMake.Builder.Document

open System.IO
open Microsoft.Office.Interop

open DocMake.Builder.BuildMonad


/// Document has a Phantom Type so we can distinguish between different types 
/// (Word, Excel, Pdf, ...)
/// Maybe we ought to store whether a file has been derived in the build process
/// (and so deletable)... 
type Document<'a> = { DocumentPath : string }

let castDocument (doc:Document<'a>) : Document<'b> = 
    { DocumentPath = doc.DocumentPath }
    


type ExcelPhantom = class end
type ExcelDoc = Document<ExcelPhantom>

type WordPhantom = class end
type WordDoc = Document<WordPhantom>

type PowerPointPhantom = class end
type PowerPointDoc = Document<PowerPointPhantom>

type PdfPhantom = class end
type PdfDoc = Document<PdfPhantom>



let makeDocument (filePath:string) : Document<'a> = 
    { DocumentPath = filePath }

let documentName (doc:Document<'a>) : string = 
    System.IO.FileInfo(doc.DocumentPath).Name


let freshDocument () : BuildMonad<'res,Document<'a>> = 
    fmapM makeDocument <| freshFileName ()

let documentExtension (doc:Document<'a>) : string = 
    System.IO.FileInfo(doc.DocumentPath).Extension


let documentDirectory (doc:Document<'a>) : string = 
    System.IO.FileInfo(doc.DocumentPath).DirectoryName


// TODO should this change assert the Phantom?
let documentChangeExtension (extension: string) (doc:Document<'a>) : Document<'b> = 
    let d1 = System.IO.Path.ChangeExtension(doc.DocumentPath, extension)
    makeDocument d1


let castToPdf (doc:Document<'a>) : PdfDoc = castDocument doc
let castToXls (doc:Document<'a>) : ExcelDoc = castDocument doc
let castToDoc (doc:Document<'a>) : WordDoc = castDocument doc
let castToPpt (doc:Document<'a>) : PowerPointDoc = castDocument doc