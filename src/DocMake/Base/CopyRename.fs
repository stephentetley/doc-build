module DocMake.Base.CopyRename

open System.IO
open System.Text.RegularExpressions


open DocMake.Base.Common

open Fake
open Fake.Core
open Fake.Core.Globbing.Operators

// TODO - this is the code from DocMake.Tasks.UniformRename.
// It should not be a task (have params record) and needs cleaning up
// to have a nice API.
 
// Rename files in a folder matching a pattern. 
// The new names are uniform base name with an numbering.
// The file extension is kept the same.

// Push whether or not to use sprintf to the client, this makes things more flexible
type NameFormatter = int->string


// Use... System.IO.File.Move (oldname,newname)
// System.IO.Directory.GetFiles(directory)

let multiCopyGlobRename  (srcDir:string, srcGlob:string) (destDir:string, destNamer:int -> string) : unit = 
    let inputs = findAllMatchingFiles srcGlob srcDir  
    List.iteri (fun ix srcFile ->
                    let destFile = destDir @@ destNamer ix
                    Fake.IO.Shell.CopyFile destFile srcFile) inputs


let private regexMatchFiles (srcDir:string) (search:string) (ignoreCase:bool) : string list = 
    let re = if ignoreCase then new Regex(search, RegexOptions.IgnoreCase) else new Regex(search)
    Directory.GetFiles(srcDir) |> Array.filter (fun s -> re.Match(s).Success) |> Array.toList



let multiCopyRegexRename  (srcDir:string, srcRegex:string, ignoreCase:bool) (destDir:string, destNamer:int -> string) : unit = 
    let inputs = regexMatchFiles srcDir srcRegex ignoreCase
    List.iteri (fun ix srcFile ->
                    let destFile = destDir @@ destNamer ix
                    Fake.IO.Shell.CopyFile destFile srcFile) inputs