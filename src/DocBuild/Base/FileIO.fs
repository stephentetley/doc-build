// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Base

[<AutoOpen>]
module FileIO = 

    open System.IO
    open System

    open DocBuild.Base.DocMonad
    open DocBuild.Base.FakeLike


    let askWorkingDirectory : DocMonad<string> = 
        asks (fun env -> env.WorkingDirectory)
    

    /// Return the full path of a filename local to the working directory.
    /// Does not validate if the file exists
    let askWorkingFile (fileName:string) : DocMonad<string> = 
        docMonad { 
            let! cwd = askWorkingDirectory
            let path = cwd </> fileName
            return path
        }

    let createWorkingSubDirectory (subDirectory:string) : DocMonad<unit> = 
        let create1 (path:string) : DocMonad<unit> = 
            if Directory.Exists(path) then
                breturn ()
            else
                Directory.CreateDirectory(path) |> ignore
                breturn ()
        docMonad {
            let! cwd = askWorkingDirectory
            let path = cwd </> subDirectory
            do! attempt (create1 path)
        }

    /// Run an operation in a subdirectory of current working directory.
    /// Create the directory if it doesn't exist.
    let localSubDirectory (subDirectory:string) 
                          (ma:DocMonad<'a>) : DocMonad<'a> = 
        docMonad {
            let! cwd = askWorkingDirectory
            let path = cwd </> subDirectory
            do! createWorkingSubDirectory path
            let! ans = local (fun env -> {env with WorkingDirectory = path}) ma
            return ans
        }

    /// This is a bit too primitive, ideally it would work on Documents.
    let copyToWorking (doc:Document<'a>) : DocMonad<Document<'a>> = 
        if File.Exists(doc.Path) then 
            docMonad { 
                let justFile = Path.GetFileName(doc.Path)
                let! cwd = askWorkingDirectory
                let target = cwd </> justFile
                do File.Copy( sourceFileName = doc.Path
                            , destFileName = target )
                return Document(target)
            }
        else
            throwError 
                <| sprintf "copyToWorking: sourceFile not found '%s'" doc.Path
            
        