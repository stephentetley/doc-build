﻿// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Base

[<AutoOpen>]
module FileIO = 

    open System.IO
    open System

    open DocBuild.Base
    open DocBuild.Base.DocMonad
    open DocBuild.Base.DocMonadOperators


    /// Note if the second path is prefixed by '\\'
    /// "directory" </> "/file.ext" == "/file.ext"
    let (</>) (path1:string) (path2:string) = 
        Path.Combine(path1, path2)


    let askWorkingDirectory () : DocMonad<'res,Uri> = 
        asks (fun env -> env.WorkingDirectory)

    let askSourceDirectory () : DocMonad<'res,Uri> = 
        asks (fun env -> env.SourceDirectory)
        
    let askIncludeDirectory () : DocMonad<'res,Uri> = 
        asks (fun env -> env.IncludeDirectory)

    let private askFile (getTopLevel:unit -> DocMonad<'res,Uri>)
                        (fileName:string) : DocMonad<'res,Uri> = 
        docMonad { 
            let! cwd = getTopLevel ()
            let path = cwd.AbsolutePath </> fileName
            return new Uri(path)
        }

    /// Return the full path of a filename local to the working directory.
    /// Does not validate if the file exists
    let askWorkingFile (fileName:string) : DocMonad<'res,Uri> = 
        askFile askWorkingDirectory fileName


    /// Return the full path of a filename local to the Source directory.
    /// Does not validate if the file exists
    let askSourceFile (fileName:string) : DocMonad<'res,Uri> = 
        askFile askSourceDirectory fileName


    /// Return the full path of a filename local to the Include directory.
    /// Does not validate if the file exists
    let askIncludeFile (fileName:string) : DocMonad<'res,Uri> = 
        askFile askIncludeDirectory fileName
    

    let createWorkingSubDirectory (subDirectory:string) : DocMonad<'res,unit> = 
        let create1 (path:string) : DocMonad<'res,unit> = 
            if Directory.Exists(path) then
                dreturn ()
            else
                Directory.CreateDirectory(path) |> ignore
                dreturn ()
        docMonad {
            let! cwd = askWorkingDirectory ()
            let path = cwd.AbsolutePath </> subDirectory
            do! attempt (create1 path)
        }

    /// Run an operation in a subdirectory of current working directory.
    /// Create the directory if it doesn't exist.
    let localSubDirectory (subDirectory:string) 
                          (ma:DocMonad<'res,'a>) : DocMonad<'res,'a> = 
        docMonad {
            let! cwd = askWorkingDirectory ()
            let path = cwd.AbsolutePath </> subDirectory
            do! createWorkingSubDirectory path
            let! ans = local (fun env -> {env with WorkingDirectory = new Uri(path)}) ma
            return ans
        }

    /// Run an operation with the Source directory restricted to the
    /// supplied sub-directory.
    let childSourceDirectory (subDirectory:string) 
                             (ma:DocMonad<'res,'a>) : DocMonad<'res,'a> = 
        docMonad {
            let! srcDir = askSourceDirectory ()
            let path = srcDir.AbsolutePath </> subDirectory
            let! ans = local (fun env -> {env with SourceDirectory = new Uri(path)}) ma
            return ans
        }

    /// This will overwrite existing documents!
    let copyToWorking (doc:Document<'a>) : DocMonad<'res,Document<'a>> = 
            docMonad { 
                let justFile = Path.GetFileName(doc.Path.AbsolutePath)
                let! cwd = askWorkingDirectory ()
                let target = cwd.AbsolutePath </> justFile
                do if File.Exists(target) then File.Delete(target) else ()
                do File.Copy( sourceFileName = doc.Path.AbsolutePath
                            , destFileName = target )
                return Document(target)
            }
  


    let copyCollectionToWorking (col:Collection.Collection<'a>) : DocMonad<'res, Collection.Collection<'a>> = 
        Collection.mapM copyToWorking col


    /// Change to internal file path to point to the working directory.
    /// This does not physically copy the file.
    let changeToWorkingFile (fileName:string) : DocMonad<'res,Uri> = 
        docMonad { 
            let! cwd = askWorkingDirectory ()
            let path = cwd.AbsolutePath </> fileName
            return new Uri(path)
        }            

    // ************************************************************************
    // Source files
            

    /// Has one or more matches. 
    /// Note - pattern is a simple glob 
    /// (the only wild cards are '?' and '*'), not a regex.
    let hasSourceFilesMatching (pattern:string) : DocMonad<'res, bool> = 
        askSourceDirectory () |>>  fun uri -> 
            FakeLike.hasFilesMatching pattern uri.AbsolutePath

    /// Search file matching files in the SourceDirectory.
    /// Uses glob pattern - the only wild cards are '?' and '*'
    let findAllSourceFilesMatching (pattern:string) : DocMonad<'res, string list> =
        askSourceDirectory () |>> fun uri -> 
            FakeLike.findAllFilesMatching pattern uri.AbsolutePath
            