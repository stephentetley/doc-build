module DocMake.Base.CopyRename

open System.IO
open System.Text.RegularExpressions



open Fake
open Fake.Core
open Fake.Core.Globbing.Operators


let private regexMatchFiles (srcDir:string) (search:string) (ignoreCase:bool) : string list = 
    let re = if ignoreCase then new Regex(search, RegexOptions.IgnoreCase) else new Regex(search)
    Directory.GetFiles(srcDir) |> Array.filter (fun s -> re.Match(s).Success) |> Array.toList




let multiCopyGlob  (srcDir:string, srcGlob:string) (destDir:string) : unit = 
    let inputs = findAllMatchingFiles srcGlob srcDir  
    List.iter (fun srcFile ->
                    Fake.IO.Shell.CopyFile destDir srcFile) inputs


let multiCopyRegex (srcDir:string, srcRegex:string, ignoreCase:bool) (destDir:string) : unit = 
    let inputs = regexMatchFiles srcDir srcRegex ignoreCase
    List.iter (fun srcFile ->
                    Fake.IO.Shell.CopyFile destDir srcFile) inputs

// Push whether or not to use sprintf to the client, this makes things 
// easier for the API.
type NameFormatter = int->string


let multiCopyGlobRename  (srcDir:string, srcGlob:string) (destDir:string, destNamer:int -> string) : unit = 
    let inputs = findAllMatchingFiles srcGlob srcDir  
    List.iteri (fun ix srcFile ->
                    let destFile = destDir @@ destNamer ix
                    Fake.IO.Shell.CopyFile destFile srcFile) inputs



let multiCopyRegexRename  (srcDir:string, srcRegex:string, ignoreCase:bool) (destDir:string, destNamer:int -> string) : unit = 
    let inputs = regexMatchFiles srcDir srcRegex ignoreCase
    List.iteri (fun ix srcFile ->
                    let destFile = destDir @@ destNamer ix
                    Fake.IO.Shell.CopyFile destFile srcFile) inputs

// Throws error if the source is not found...
let mandatoryCopyFile (destPath:string) (source:string) : unit = 
    if IO.File.exists(source) then
        Fake.IO.Shell.CopyFile destPath source
    else 
        failwithf "mandatoryCopyFile - source not found '%s'" source

// Prints warning if the source is not found...
let optionalCopyFile (destPath:string) (source:string) : unit = 
    if IO.File.exists(source) then
        Fake.IO.Shell.CopyFile destPath source
    else 
        Trace.tracefn "optionalCopyFile: WARNING - not copied, source not found '%s'" source
