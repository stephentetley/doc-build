// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

namespace DocBuild.Internal.CommonUtils


[<AutoOpen>]
module CommonUtils = 

    open System.IO
    open System.Text.RegularExpressions

    let doubleQuote (s:string) : string = "\"" + s + "\""


    let rbox (x:'a) : ref<obj> = ref (x :> obj)


    /// Make a file name for "destructive update".
    /// Returns a file name with ``DBU`` before the extension.
    /// If the file already contains with ``DBU`` before the extension, 
    /// return the file name so it can be overwritten.
    let getTempFileName (filePath:string) : string = 
        let justfile = System.IO.Path.GetFileNameWithoutExtension filePath
        if Regex.Match(justfile, "\.DBU$").Success then 
            filePath
        else
            let root = System.IO.Path.GetDirectoryName filePath
            let ext  = System.IO.Path.GetExtension filePath
            let newfile = sprintf "%s.DBU.%s" justfile ext
            Path.Combine(root, newfile)

