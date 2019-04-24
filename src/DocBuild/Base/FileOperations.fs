// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Base

[<AutoOpen>]
module FileOperations = 

    open System.IO
    open System

    open DocBuild.Base.Internal
    open DocBuild.Base
    open DocBuild.Base.DocMonad

    /// TODO - idea of working directory is too complicated.
    /// And we need a simple way to make temp files.


    let extendWorkingPath (relPath:string) : DocMonad<'userRes,string> = 
        askWorkingDirectory () |>> fun root -> root </> relPath

    let extendSourcePath (relPath:string) : DocMonad<'userRes,string> = 
        askSourceDirectory () |>> fun root -> root </> relPath
    


    let isWorkingPath (absPath:string) : DocMonad<'userRes,bool> = 
        askWorkingDirectory () |>> fun dir -> FilePaths.rootIsPrefix dir absPath


    let isWorkingDocument (doc:Document<'a>) : DocMonad<'userRes,bool> = 
        isWorkingPath doc.AbsolutePath

    let isSourcePath (absPath:string) : DocMonad<'userRes,bool> = 
        askSourceDirectory () |>> fun dir -> FilePaths.rootIsPrefix dir absPath
        

    let isSourceDocument (doc:Document<'a>) : DocMonad<'userRes,bool> = 
        isSourcePath doc.AbsolutePath


    let isIncludePath (absPath:string) : DocMonad<'userRes,bool> = 
        askIncludeDirectories () |>> fun dirs -> 
        List.exists (fun dir -> FilePaths.rootIsPrefix dir absPath) dirs 
        
    
    let isIncludeDocument (doc:Document<'a>) : DocMonad<'userRes,bool> = 
        isIncludePath doc.AbsolutePath



    let assertIsWorkingPath (path:string) : DocMonad<'userRes, unit> = 
        assertM (sprintf "Not a working path - '%s'" path)
                (isWorkingPath path)

    let assertIsWorkingDocument (doc:Document<'a>) : DocMonad<'userRes, unit> = 
        assertM (sprintf "Not a working Document - '%s'" doc.Title)
                (isWorkingDocument doc)

    let assertIsSourcePath (path:string) : DocMonad<'userRes, unit> = 
        assertM (sprintf "Not a source path - '%s'" path)
                (isSourcePath path)

    let assertIsSourceDocument (doc:Document<'a>) : DocMonad<'userRes, unit> = 
        assertM (sprintf "Not a source Document - '%s'" doc.Title)
                (isSourceDocument doc) 


    let assertIsIncludePath (path:string) : DocMonad<'userRes, unit> = 
        assertM (sprintf "Not an include path - '%s'" path)
                (isIncludePath path)

    let assertIsIncludeDocument (doc:Document<'a>) : DocMonad<'userRes, unit> = 
        assertM (sprintf "Not an include Document - '%s'" doc.Title)
                (isIncludeDocument doc) 

    let getWorkingPathSuffix (absPath:string) : DocMonad<'userRes,string> = 
        docMonad { 
            do! assertIsWorkingPath absPath
            let! root = askWorkingDirectory ()
            return FilePaths.rightPathComplement root absPath
        }
            
    let getWorkingDocPathSuffix (doc:Document<'a>) : DocMonad<'userRes,string> = 
        getWorkingPathSuffix doc.AbsolutePath


    let getSourcePathSuffix (absPath:string) : DocMonad<'userRes,string> = 
        docMonad { 
            do! assertIsSourcePath absPath
            let! root = askSourceDirectory ()
            return FilePaths.rightPathComplement root absPath
        }

    let getSourceDocPathSuffix (doc:Document<'a>) : DocMonad<'userRes,string> = 
        getSourcePathSuffix doc.AbsolutePath

    //let getIncludePathSuffix (absPath:string) : DocMonad<'userRes,string> = 
    //    docMonad { 
    //        do! assertIsIncludePath absPath
    //        let! root = askIncludeDirectory ()
    //        return rightPathComplement root (FilePath(absPath))
    //    }

    //let getIncludeDocPathSuffix (doc:Document<'a>) : DocMonad<'userRes,string> = 
    //    getIncludePathSuffix doc.LocalPath

    //let getPathSuffix (absPath:string) : DocMonad<'userRes,string> =
    //    getWorkingPathSuffix absPath 
    //        <||> getSourcePathSuffix absPath 
    //        <||> getIncludePathSuffix absPath 

    //let getDocPathSuffix (doc:Document<'a>) : DocMonad<'userRes,string> =
    //    getWorkingDocPathSuffix doc 
    //        <||> getSourceDocPathSuffix doc
    //        <||> getIncludeDocPathSuffix doc

    /// Create a subdirectory in the Working directory.
    let createWorkingFolder (subfolderName:string) : DocMonad<'userRes,unit> = 
        askWorkingDirectory () |>> fun cwd ->
        let absFolderName = cwd </> subfolderName
        if Directory.Exists(absFolderName) then
            ()
        else Directory.CreateDirectory(absFolderName) |> ignore
            

    ///// Rewrite the the file name to site it in the working folder.
    ///// If the file is from Source or Include directories generate the name with 
    ///// the respective subfolder path from root.
    ///// Otherwise, generate the file name at the top level of Workgin.
    //let generateWorkingFileName (includeDirectoriesSuffix:bool) (absPath:string) : DocMonad<'userRes,string> = 
    //    docMonad { 
    //        let! suffix = 
    //            if includeDirectoriesSuffix then 
    //                getPathSuffix absPath <||> mreturn (FileInfo(absPath).Name)
    //            else mreturn (FileInfo(absPath).Name)
    //        return! extendWorkingPath suffix
    //    }

    /// Copy a file to working, returning the copy as a Document.
    /// If the file is from Source or Include directories copy with 
    /// the respective subfolder path from root.
    let private copyFileToWorking (sourceAbsPath:string) : DocMonad<'userRes,Document<'a>> = 
        docMonad { 
            let name = FileInfo(sourceAbsPath).Name
            let! target = extendWorkingPath name
            if File.Exists(target) then 
                File.Delete(target) 
            else ()
            File.Copy( sourceFileName = sourceAbsPath, destFileName = target )
            return Document(target)
        }

    /// Copy a doc to the toplevel working directory.
    let copyToWorking (doc:Document<'a>) : DocMonad<'userRes,Document<'a>> = 
        let title = doc.Title
        copyFileToWorking doc.AbsolutePath |>> setTitle title




    /// Rename a folder in the working drectory
    let renameWorkingFolder (oldName:string) (newName:string) : DocMonad<'userRes,unit> = 
        askWorkingDirectory () >>= fun cwd ->
        let oldPath = cwd </> oldName
        if Directory.Exists(oldPath) then
            let newPath = cwd </> newName
            Directory.Move(oldPath, newPath)
            mreturn ()
        else
            throwError (sprintf "renameWorkingFolder - folder does not exist '%s'" oldPath)

    /// Has one or more matches. 
    /// Note - pattern is a simple glob 
    /// (the only wild cards are '?' and '*'), not a regex.
    let hasSourceFilesMatching (pattern:string) 
                               (recurseIntoSubDirectories:bool) : DocMonad<'userRes, bool> = 
        let proc () = 
            askSourceDirectory () |>>  FakeLikePrim.hasFilesMatching pattern recurseIntoSubDirectories
        attemptM <| proc ()



 
    /// Search file matching files in the SourceDirectory.
    /// Uses glob pattern - the only wild cards are '?' and '*'
    /// Returns a list of absolute paths.
    let findAllSourceFilesMatching (pattern:string) 
                                   (recurseIntoSubDirectories:bool) : DocMonad<'userRes, string list> =
        let proc () = 
            askSourceDirectory () |>> FakeLikePrim.findAllFilesMatching pattern recurseIntoSubDirectories
        attemptM <| proc ()

    /// Search file matching files in the SourceDirectory.
    /// Uses glob pattern - the only wild cards are '?' and '*'
    /// Returns a list of absolute paths.
    let findSomeSourceFilesMatching (pattern:string) 
                                    (recurseIntoSubDirectories:bool) : DocMonad<'userRes, string list> =
        let proc () = 
            askSourceDirectory () |>> FakeLikePrim.tryFindSomeFilesMatching pattern recurseIntoSubDirectories
        optionFailM "fail - findSomeSourceFilesMatching" <| proc ()

    /// Search file matching files in the SourceDirectory.
    /// Uses glob pattern - the only wild cards are '?' and '*'
    /// Returns a list of absolute paths.
    let tryFindExactlyOneSourceFileMatching (pattern:string) 
                                   (recurseIntoSubDirectories:bool) : DocMonad<'userRes, string option> =
        let proc () = 
            askSourceDirectory () |>> FakeLikePrim.tryFindExactlyOneFileMatching pattern recurseIntoSubDirectories
        attemptM <| proc ()


    /// Create a subdirectory under the working folder.
    let createWorkingSubdirectory (relPath:string) : DocMonad<'userRes,unit> = 
        askWorkingDirectory () |>> fun cwd ->
        let path = cwd </> relPath 
        if Directory.Exists(path) then 
            ()
        else Directory.CreateDirectory(path) |> ignore


    /// Run an operation in a subdirectory of current working directory.
    /// Creates the subdirectory if it doesn't exist.
    let localWorkingSubdirectory (subdirectory:string) 
                                 (ma:DocMonad<'userRes,'a>) : DocMonad<'userRes,'a> = 
        docMonad {
            let! path = extendWorkingPath subdirectory
            let! _ = createWorkingSubdirectory subdirectory
            return! local (fun env -> {env with WorkingDirectory = path}) ma
        }
            
    /// Run an operation with the Source directory restricted to the
    /// supplied subdirectory.
    let localSourceSubdirectory (subdirectory:string) 
                                (ma:DocMonad<'userRes,'a>) : DocMonad<'userRes,'a> = 
        docMonad {
            let! path = extendSourcePath subdirectory
            return! local (fun env -> {env with SourceDirectory = path}) ma
        }


