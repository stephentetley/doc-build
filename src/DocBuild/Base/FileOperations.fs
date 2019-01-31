// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Base

[<AutoOpen>]
module FileOperations = 

    open System.IO
    open System

    open DocBuild.Base
    open DocBuild.Base.DocMonad
    open DocBuild.Base.DocMonadOperators

    // This module is becoming obsolete as we are approaching a strong
    // notion of "source", "working" and "include" directories and 
    // reneging on the idea of random access to the file system.



    let getOutputPath (relativeFileName:string) : DocMonad<'res,string> = 
        askWorkingDirectory () |>> fun cwd -> (cwd.LocalPath </> relativeFileName)


    let askWorkingDirectoryPath () : DocMonad<'res,string> = 
        askWorkingDirectory () |>> fun uri -> uri.LocalPath

    let askSourceDirectoryPath () : DocMonad<'res,string> = 
        askSourceDirectory () |>> fun uri -> uri.LocalPath

    let askIncludeDirectoryPath () : DocMonad<'res,string> = 
        askIncludeDirectory () |>> fun uri -> uri.LocalPath

    let isWorkingPath (absPath:string) : DocMonad<'res,bool> = 
        docMonad { 
            let! dir = askWorkingDirectory ()
            return dir.IsBaseOf(new Uri(absPath))
        }

    let isWorkingDocument (doc:Document<'a>) : DocMonad<'res,bool> = 
        isWorkingPath doc.LocalPath

    let isSourcePath (absPath:string) : DocMonad<'res,bool> = 
        docMonad { 
            let! dir = askSourceDirectory ()
            return dir.IsBaseOf(new Uri(absPath))
        }

    let isSourceDocument (doc:Document<'a>) : DocMonad<'res,bool> = 
        isSourcePath doc.LocalPath


    let isIncludePath (absPath:string) : DocMonad<'res,bool> = 
        docMonad { 
            let! dir = askIncludeDirectory ()
            return dir.IsBaseOf(new Uri(absPath))
        }
    
    let isIncludeDocument (doc:Document<'a>) : DocMonad<'res,bool> = 
        isIncludePath doc.LocalPath

    let assertIsWorkingPath (path:string) : DocMonad<'res, unit> = 
        isWorkingPath path >>= fun ans ->
        if ans then 
            dreturn ()
        else
            throwError (sprintf "Not a working path - '%s'" path)

    let assertIsWorkingDocument (doc:Document<'a>) : DocMonad<'res, unit> = 
        isWorkingDocument doc >>= fun ans ->
        if ans then 
            dreturn ()
        else
            throwError (sprintf "Not a working Document - '%s'" doc.Title)

    let assertIsSourcePath (path:string) : DocMonad<'res, unit> = 
        isSourcePath path >>= fun ans ->
        if ans then 
            dreturn ()
        else
            throwError (sprintf "Not a source path - '%s'" path)

    let assertIsSourceDocument (doc:Document<'a>) : DocMonad<'res, unit> = 
        isSourceDocument doc >>= fun ans ->
        if ans then 
            dreturn ()
        else
            throwError (sprintf "Not a source Document - '%s'" doc.Title)


    let assertIsIncludePath (path:string) : DocMonad<'res, unit> = 
        isIncludePath path >>= fun ans ->
        if ans then 
            dreturn ()
        else
            throwError (sprintf "Not an include path - '%s'" path)

    let assertIsIncludeDocument (doc:Document<'a>) : DocMonad<'res, unit> = 
        isIncludeDocument doc >>= fun ans ->
        if ans then 
            dreturn ()
        else
            throwError (sprintf "Not an include Document - '%s'" doc.Title)

    let getWorkingPathSuffix (absPath:string) : DocMonad<'res,string> = 
        assertIsWorkingPath absPath >>= fun _ -> 
        docMonad { 
            let! dir = askWorkingDirectory ()
            let uri = dir.MakeRelativeUri(new Uri(absPath))
            return uri.ToString()
        }
            
    let getWorkingDocPathSuffix (doc:Document<'a>) : DocMonad<'res,string> = 
        getWorkingPathSuffix doc.LocalPath


    let getSourcePathSuffix (absPath:string) : DocMonad<'res,string> = 
        assertIsSourcePath absPath >>= fun _ -> 
        docMonad { 
            let! dir = askSourceDirectory ()
            let uri = dir.MakeRelativeUri(new Uri(absPath))
            return uri.ToString()
        }

    let getSourceDocPathSuffix (doc:Document<'a>) : DocMonad<'res,string> = 
        getSourcePathSuffix doc.LocalPath

    let getIncludePathSuffix (absPath:string) : DocMonad<'res,string> = 
        assertIsIncludePath absPath >>= fun _ -> 
        docMonad { 
            let! dir = askIncludeDirectory ()
            let uri = dir.MakeRelativeUri(new Uri(absPath))
            return uri.ToString()
        }

    let getIncludeDocPathSuffix (doc:Document<'a>) : DocMonad<'res,string> = 
        getIncludePathSuffix doc.LocalPath

    let getPathSuffix (path:string) : DocMonad<'res,string> =
        getWorkingPathSuffix path 
            <||> getSourcePathSuffix path 
            <||> getIncludePathSuffix path 

    let getDocPathSuffix (doc:Document<'a>) : DocMonad<'res,string> =
        getWorkingDocPathSuffix doc 
            <||> getSourceDocPathSuffix doc
            <||> getIncludeDocPathSuffix doc

    /// Create a subdirectory in the Working directory.
    let createWorkingFolder (subfolderName:string) : DocMonad<'res,unit> = 
        askWorkingDirectoryPath () >>= fun cwd ->
        let absFolderName = cwd </> subfolderName
        if Directory.Exists(absFolderName) then
            dreturn ()
        else
            Directory.CreateDirectory(absFolderName) |> ignore
            dreturn ()

    /// Copy a file to working, returnin the copy as a Document.
    /// If the file is from Source or Include directories copy with 
    /// the respective subfolder path from root.
    let copyFileToWorking (path:string) : DocMonad<'res,Document<'a>> = 
        docMonad { 
            let! suffix = getPathSuffix path <||> dreturn (FileInfo(path).Name)
            let! target = askWorkingDirectory () |>> fun uri -> uri <//> suffix
            do if File.Exists(target.LocalPath) then 
                    File.Delete(target.LocalPath) 
               else ()
            do File.Copy( sourceFileName = path
                        , destFileName = target.LocalPath )
            return Document(target)
        }

    /// Copy a doc to working.
    /// If the file is from Source or Include copy with the respective
    /// subfolder path from root.
    let copyToWorking (doc:Document<'a>) : DocMonad<'res,Document<'a>> = 
        copyFileToWorking doc.LocalPath


    /// Rename a folder in the working drectory
    let renameWorkingFolder (oldName:string) (newName:string) : DocMonad<'res,unit> = 
        askWorkingDirectory () >>= fun cwd ->
        let oldPath = (cwd <//> oldName).LocalPath
        if Directory.Exists(oldPath) then
            let newPath = (cwd <//> newName).LocalPath
            Directory.Move(oldPath, newPath)
            dreturn ()
        else
            throwError (sprintf "renameWorkingFolder - folder does not exist '%s'" oldPath)

    /// Has one or more matches. 
    /// Note - pattern is a simple glob 
    /// (the only wild cards are '?' and '*'), not a regex.
    let hasSourceFilesMatching (pattern:string) 
                               (recurseIntoSubDirectories:bool) : DocMonad<'res, bool> = 
        askSourceDirectoryPath () |>>  fun path -> 
            FakeLikePrim.hasFilesMatching pattern recurseIntoSubDirectories path




    /// Throws error if the doc to be modified is not in the working 
    /// directory.
    let withWorkingDoc (modify:Document<'a> -> DocMonad<'res,'answer>) 
                       (doc:Document<'a>) : DocMonad<'res,'answer> = 
        isWorkingDocument doc >>= fun ans -> 
        match ans with
        | true -> modify doc
        | false -> 
            throwError (sprintf "Document '%s' is not in the Working directory" doc.Title)


    /// Search file matching files in the SourceDirectory.
    /// Uses glob pattern - the only wild cards are '?' and '*'
    let findAllSourceFilesMatching (pattern:string) 
                                   (recurseIntoSubDirectories:bool) : DocMonad<'res, string list> =
        docMonad { 
            let! srcPath = askSourceDirectoryPath () 
            let fullPaths = FakeLikePrim.findAllFilesMatching pattern recurseIntoSubDirectories srcPath
            let! relPaths = mapM getSourcePathSuffix fullPaths
            return relPaths
        }

            


    //// Old API...

    let private askFile (getTopLevel:unit -> DocMonad<'res,Uri>)
                        (fileName:string) : DocMonad<'res,Uri> = 
        docMonad { 
            let! cwd = getTopLevel ()
            let path = cwd.LocalPath </> fileName
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
            let path = cwd.LocalPath </> subDirectory
            do! attempt (create1 path)
        }

    /// Run an operation in a subdirectory of current working directory.
    /// Create the directory if it doesn't exist.
    let localSubDirectory (subDirectory:string) 
                          (ma:DocMonad<'res,'a>) : DocMonad<'res,'a> = 
        docMonad {
            let! cwd = askWorkingDirectory ()
            let path = cwd.LocalPath </> subDirectory
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
            let path = srcDir.LocalPath </> subDirectory
            let! ans = local (fun env -> {env with SourceDirectory = new Uri(path)}) ma
            return ans
        }


  


    let copyCollectionToWorking (col:Collection.Collection<'a>) : DocMonad<'res, Collection.Collection<'a>> = 
        Collection.mapM copyToWorking col


    /// Change to internal file path to point to the working directory.
    /// This does not physically copy the file.
    let changeToWorkingFile (fileName:string) : DocMonad<'res,Uri> = 
        docMonad { 
            let! cwd = askWorkingDirectory ()
            let path = cwd.LocalPath </> fileName
            return new Uri(path)
        }            

    // ************************************************************************
    // Source files
            


    ///// Search file matching files in the SourceDirectory.
    ///// Uses glob pattern - the only wild cards are '?' and '*'
    //let findAllSourceFilesMatching (pattern:string) 
    //                               (recurseIntoSubDirectories:bool) : DocMonad<'res, string list> =
    //    askSourceDirectory () |>> fun uri -> 
    //        FakeLikePrim.findAllFilesMatching pattern recurseIntoSubDirectories uri.LocalPath
            