// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Base.Document

[<AutoOpen>]
module Document = 

    open System.Text.RegularExpressions
    open System.IO

    type FilePath = string
    
    /// Make a file name for "destructive update".
    /// Returns a file name with ``TEMP`` before the extension.
    /// If the file already contains with ``TEMP`` before the extension, 
    /// return the file name so it can be overwritten.
    let private getTempFileName (filePath:string) : FilePath = 
        let justfile = System.IO.Path.GetFileNameWithoutExtension filePath
        if Regex.Match(justfile, "\.TEMP$").Success then 
            filePath
        else
            let root = System.IO.Path.GetDirectoryName filePath
            let ext  = System.IO.Path.GetExtension filePath
            let newfile = sprintf "%s.TEMP.%s" justfile ext
            Path.Combine(root, newfile)


    type Document = 
        val private Original : FilePath
        val private WorkingFile : FilePath

        new (filePath:string) = 
            { Original = filePath
            ; WorkingFile = getTempFileName filePath }

        /// ActiveFile is a mutable working copy of the original file.
        /// The original file is untouched.
        member v.ActiveFile
            with get() : FilePath = 
                if System.IO.File.Exists(v.WorkingFile) then
                    v.WorkingFile
                else
                    System.IO.File.Copy(v.Original, v.WorkingFile)
                    v.WorkingFile
    
        member v.Updated 
            with get() : bool = System.IO.File.Exists(v.WorkingFile)


        member v.SaveAs(outputPath: string) : unit = 
            if v.Updated then 
                System.IO.File.Move(v.WorkingFile, outputPath)
            else
                System.IO.File.Copy(v.WorkingFile, outputPath)



