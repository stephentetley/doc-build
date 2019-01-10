// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Base

[<AutoOpen>]
module FakeLike = 

    open System.IO
    open System

    open DocBuild.Base.DocMonad
    open DocBuild.Base.FakeLike

    let validateFile (fileExtensions:string list) (path:string) : DocMonad<string> = 
        if System.IO.File.Exists(path) then 
            let extension : string = System.IO.Path.GetExtension(path)
            let testExtension (ext:string) : bool = String.Equals(extension, ext, StringComparison.CurrentCultureIgnoreCase)
            if List.exists testExtension fileExtensions then 
                breturn path
            else throwError <| sprintf "Not a %s file: '%s'" (String.concat "," fileExtensions) path
        else throwError <| sprintf "Could not find file: '%s'" path  


    /// This is a bit too primitive, ideally it would work on Documents.
    let copyToWorking (sourceFile:string) : DocMonad<string> = 
        if File.Exists(sourceFile) then 
            docMonad { 
                let justFile = Path.GetFileName(sourceFile)
                let! cwd = askWorkingDirectory
                let target = cwd </> justFile
                do File.Copy( sourceFileName = sourceFile
                            , destFileName = target )
                return target
            }
        else
            throwError <| sprintf "copyToWorking: sourceFile not found '%s'" sourceFile
            
        