// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Base


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
        val private SourcePath : FilePath
        val private TempPath : string

        new (filePath:string) = 
            { SourcePath = filePath
            ; TempPath = getTempFileName filePath }

        /// Maybe `FileName` would be better?
        member v.TempFile
            with get() : string = 
                if System.IO.File.Exists(v.TempPath) then
                    v.TempPath
                else
                    System.IO.File.Copy(v.SourcePath, v.TempPath)
                    v.TempPath
    
        member v.Updated 
            with get() : bool = System.IO.File.Exists(v.TempPath)


        member v.SaveAs(outputPath: string) : unit = 
            if v.Updated then 
                System.IO.File.Move(v.TempPath, outputPath)
            else
                System.IO.File.Copy(v.SourcePath, outputPath)


    type Aggregate = 
        val private Documents : Document list

        new () = { Documents = [] }

        new (paths:FilePath list) = 
            { Documents = 
                paths |> List.map (fun s -> new Document(filePath = s)) }


