// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocBuild.Base


[<RequireQualifiedAccess>]
module Temp = 

    open System.IO
    open System.Text.RegularExpressions

    /// Make a file name for "destructive update".
    /// Returns a file name with ``TEMP`` before the extension.
    /// If the file already contains with ``TEMP`` before the extension, 
    /// return the file name so it can be overwritten.
    let getTempFileName (filePath:string) : string = 
        let justfile = System.IO.Path.GetFileNameWithoutExtension filePath
        if Regex.Match(justfile, "\.TEMP$").Success then 
            filePath
        else
            let root = System.IO.Path.GetDirectoryName filePath
            let ext  = System.IO.Path.GetExtension filePath
            let newfile = sprintf "%s.TEMP.%s" justfile ext
            Path.Combine(root, newfile)

