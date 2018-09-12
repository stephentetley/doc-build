// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

module DocMake.Base.CopyRename

open System.IO
open System.Text.RegularExpressions

open DocMake.Base.FakeLike



let private regexMatchFiles (srcDir:string) (search:string) (ignoreCase:bool) : string list = 
    let re = if ignoreCase then new Regex(search, RegexOptions.IgnoreCase) else new Regex(search)
    Directory.GetFiles(srcDir) 
        |> Array.filter (fun s -> re.Match(s).Success) 
        |> Array.toList




let multiCopyGlob  (srcDir:string, srcGlob:string) (destDir:string) : unit = 
    let inputs = findAllMatchingFiles srcGlob srcDir  
    List.iter (fun srcFile -> copyFile destDir srcFile) inputs


let multiCopyRegex (srcDir:string, srcRegex:string, ignoreCase:bool) (destDir:string) : unit = 
    let inputs = regexMatchFiles srcDir srcRegex ignoreCase
    List.iter (fun srcFile ->
                    copyFile destDir srcFile) inputs

// Push whether or not to use sprintf to the client, this makes things 
// easier for the API.
type NameFormatter = int->string


let multiCopyGlobRename  (srcDir:string, srcGlob:string) (destDir:string, destNamer:int -> string) : unit = 
    let inputs = findAllMatchingFiles srcGlob srcDir  
    List.iteri (fun ix srcFile ->
                    let destFile = destDir </> destNamer ix
                    copyFile destFile srcFile) inputs



let multiCopyRegexRename  (srcDir:string, srcRegex:string, ignoreCase:bool) (destDir:string, destNamer:int -> string) : unit = 
    let inputs = regexMatchFiles srcDir srcRegex ignoreCase
    List.iteri (fun ix srcFile ->
                    let destFile = destDir </> destNamer ix
                    copyFile destFile srcFile) inputs

// Throws error if the source is not found...
let mandatoryCopyFile (destPath:string) (source:string) : unit = 
    if fileExists(source) then
        copyFile destPath source
    else 
        failwithf "mandatoryCopyFile - source not found '%s'" source

// Prints warning if the source is not found...
let optionalCopyFile (destPath:string) (source:string) : unit = 
    if fileExists(source) then
        copyFile destPath source
    else 
        printfn "optionalCopyFile: WARNING - not copied, source not found '%s'" source
