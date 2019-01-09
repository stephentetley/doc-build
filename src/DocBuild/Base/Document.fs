// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Base.Document

[<AutoOpen>]
module Document = 

    open System.Text.RegularExpressions
    open System.IO

    open DocBuild.Base.Monad
    open DocBuild.Base

    type FilePath = string

    /// The temp indicator is a suffix "Z0.." before the file extension
    let internal getNextTempName (filePath:FilePath) : FilePath =
        let root = System.IO.Path.GetDirectoryName filePath
        let justFile = Path.GetFileNameWithoutExtension filePath
        let extension  = System.IO.Path.GetExtension filePath

        let patt = @"Z(\d+)$"
        let result = Regex.Match(justFile, patt)
        let count = 
            if result.Success then 
                int <| result.Groups.Item(1).Value
            else 0
        let suffix = sprintf "Z%03d" (count+1)
        let newfile = sprintf "%s.%s%s" justFile suffix extension
        Path.Combine(root, newfile)

    
    [<Struct>]
    type Document = 
        | Document of FilePath 

        member x.Path 
            with get () : FilePath =
                match x with | Document(p) -> p

        /// ActiveFile is a mutable working copy of the original file.
        /// The original file is untouched.
        member x.NextTempName
            with get() : FilePath = 
                getNextTempName x.Path



    let getDocument (fileExtension:string) (filePath:string) : DocBuild<Document> = 
        docBuild { 
            let! path = validateFile fileExtension filePath
            return Document(path)
            }
