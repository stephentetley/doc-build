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
            let! cwd = asks (fun env -> env.WorkingDirectory)
            let path = cwd </> fileName
            return path
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
            
        