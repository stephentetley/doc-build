// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


module DocMake.Builder.Document

open System.IO
open Microsoft.Office.Interop

open DocMake.Builder.BuildMonad


/// Document has a Phantom Type so we can distinguish between different types 
/// (Word, Excel, Pdf, ...)
/// Document can be empty - an optional build task will generate an empty doc.
type Document<'a> = 
    private { DocumentPath : option<string> }
    member v.GetPath : option<string> = v.DocumentPath


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

type MarkdownPhantom = class end
type MarkdownDoc = Document<MarkdownPhantom>

let private mapDocumentPath (fn:string -> 'ans) (doc:Document<'a>) : option<'ans> = 
    Option.map fn doc.DocumentPath

let makeDocument (filePath:string) : Document<'a> = 
    { DocumentPath = Some filePath }

let zeroDocument : Document<'a> =  
    { DocumentPath = None } 



let documentName (doc:Document<'a>) : option<string> = 
    mapDocumentPath (fun path -> System.IO.FileInfo(path).Name) doc

let documentExtension (doc:Document<'a>) : option<string> = 
    mapDocumentPath (fun path -> System.IO.FileInfo(path).Extension) doc


let documentDirectory (doc:Document<'a>) : option<string> = 
    mapDocumentPath (fun path -> System.IO.FileInfo(path).DirectoryName) doc



let documentChangeExtension (extension: string) (doc:Document<'a>) : Document<'b> = 
    match doc.DocumentPath with
    | Some filepath -> System.IO.Path.ChangeExtension(filepath, extension) |> makeDocument
    | None -> { DocumentPath = None }


let freshDocument (extension:string) : BuildMonad<'res,Document<'a>> = 
    fmapM makeDocument <| freshFileName extension 


let workingDocument (docName:string) : BuildMonad<'res,Document<'a>> = 
    buildMonad { 
        let! cwd = askEnv () |>> (fun env -> env.WorkingDirectory)
        let outPath = System.IO.Path.Combine(cwd,docName)
        return (makeDocument outPath)
    }










//let castToPdf (doc:Document<'a>) : PdfDoc = castDocument doc
//let castToXls (doc:Document<'a>) : ExcelDoc = castDocument doc
//let castToDoc (doc:Document<'a>) : WordDoc = castDocument doc
//let castToPpt (doc:Document<'a>) : PowerPointDoc = castDocument doc