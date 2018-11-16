// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


module DocBuild.Common

open System.IO

type SearchList = (string * string) list



/// Suffix a file name _before_ the extension.
///
/// e.g suffixFileName "TEMP"  "sunset.jpg" ==> "sunset.TEMP.jpg"
let suffixFileName (suffix:string)  (filePath:string) : string = 
    let root = System.IO.Path.GetDirectoryName filePath
    let justfile = System.IO.Path.GetFileNameWithoutExtension filePath
    let ext  = System.IO.Path.GetExtension filePath
    let newfile = sprintf "%s.%s%s" justfile suffix ext
    Path.Combine(root, newfile)

