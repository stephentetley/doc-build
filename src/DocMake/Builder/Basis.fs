﻿module DocMake.Builder.Basis

open System.IO

open DocMake.Builder.BuildMonad


/// Document has a Phantom Type so we can distinguish between different types 
/// (Word, Excel, Pdf, ...)
/// Maybe we ought to store whether a file has been derived in the build process
/// (and so deleteable)... 
type Document<'a> = { DocumentPath : string }

let makeDocument (filePath:string) : Document<'a> = 
    { DocumentPath = filePath }

let freshDocument () : BuildMonad<'res,Document<'a>> = 
    fmapM makeDocument <| freshFileName ()

let documentExtension (doc:Document<'a>) : string = 
    System.IO.FileInfo(doc.DocumentPath).Extension


let documentDirectory (doc:Document<'a>) : string = 
    System.IO.FileInfo(doc.DocumentPath).DirectoryName


let documentName (doc:Document<'a>) : string = 
    System.IO.FileInfo(doc.DocumentPath).Name


let assertFile(fileName:string) : BuildMonad<'res,string> =  
    if File.Exists(fileName) then 
        buildMonad.Return(fileName)
    else 
        throwError <| sprintf "assertFile failed: '%s'" fileName

// assertExtension ?

let getDocument(fileName:string) : BuildMonad<'res,Document<'a>> =   
    if File.Exists(fileName) then 
        buildMonad.Return({DocumentPath=fileName})
    else 
        throwError <| sprintf "getDocument failed: '%s'" fileName


let renameDocument (src:Document<'a>) (dest:string) : BuildMonad<'res,Document<'a>> =  
    executeIO <| fun () -> 
        let srcPath = src.DocumentPath
        let pathTo = documentDirectory src
        let outPath = System.IO.Path.Combine(pathTo,dest)
        if System.IO.File.Exists(outPath) then System.IO.File.Delete(outPath)
        System.IO.File.Move(srcPath,outPath)
        {DocumentPath=outPath}

let renameTo (dest:string) (src:Document<'a>) : BuildMonad<'res,Document<'a>> = 
    renameDocument src dest