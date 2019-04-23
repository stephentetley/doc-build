// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Base

[<AutoOpen>]
module FileOperations = 

    open System.IO
    open System

    open DocBuild.Base
    open DocBuild.Base.DocMonad

    /// TODO - idea of working directory is too complicated.
    /// And we need a simple way to make temp files.

    /// Modify the file name, leave the directory path 
    /// and the extension the same.
    let modifyFileName (modify:string -> string) 
                       (absPath:string) : string = 
        let left = Path.GetDirectoryName absPath
        let name = Path.GetFileNameWithoutExtension absPath
        let name2 = 
            if Path.HasExtension absPath then
                let ext = Path.GetExtension absPath
                modify name + ext
            else
                modify name
        Path.Combine(left, name2)


    let getOutputPath (relativeFileName:string) : DocMonad<'res,string> = 
        askWorkingDirectory () |>> fun cwd -> (cwd.LocalPath </> relativeFileName)


    let askWorkingDirectoryPath () : DocMonad<'res,string> = 
        askWorkingDirectory () |>> fun path -> path.LocalPath

    let askSourceDirectoryPath () : DocMonad<'res,string> = 
        askSourceDirectory () |>> fun path -> path.LocalPath



    let extendWorkingPath (relPath:string) : DocMonad<'res,string> = 
        askWorkingDirectoryPath () |>> fun root -> root </> relPath

    let extendSourcePath (relPath:string) : DocMonad<'res,string> = 
        askSourceDirectoryPath () |>> fun root -> root </> relPath
    


    let isWorkingPath (absPath:string) : DocMonad<'res,bool> = 
        askWorkingDirectory () |>> fun dir -> rootIsPrefix dir (FilePath(absPath))


    let isWorkingDocument (doc:Document<'a>) : DocMonad<'res,bool> = 
        isWorkingPath doc.AbsolutePath

    let isSourcePath (absPath:string) : DocMonad<'res,bool> = 
        askSourceDirectory () |>> fun dir -> rootIsPrefix dir (FilePath(absPath))
        

    let isSourceDocument (doc:Document<'a>) : DocMonad<'res,bool> = 
        isSourcePath doc.AbsolutePath


    let isIncludePath (absPath:string) : DocMonad<'res,bool> = 
        askIncludeDirectories () |>> fun dirs -> 
        List.exists (fun dir -> rootIsPrefix (DirectoryPath dir) (FilePath(absPath))) dirs 
        
    
    let isIncludeDocument (doc:Document<'a>) : DocMonad<'res,bool> = 
        isIncludePath doc.AbsolutePath

    let assertIsWorkingPath (path:string) : DocMonad<'res, unit> = 
        assertM (isWorkingPath path) (sprintf "Not a working path - '%s'" path)

    let assertIsWorkingDocument (doc:Document<'a>) : DocMonad<'res, unit> = 
        assertM (isWorkingDocument doc) (sprintf "Not a working Document - '%s'" doc.Title)

    let assertIsSourcePath (path:string) : DocMonad<'res, unit> = 
        assertM (isSourcePath path) (sprintf "Not a source path - '%s'" path)

    let assertIsSourceDocument (doc:Document<'a>) : DocMonad<'res, unit> = 
        assertM (isSourceDocument doc) (sprintf "Not a source Document - '%s'" doc.Title)


    let assertIsIncludePath (path:string) : DocMonad<'res, unit> = 
        assertM (isIncludePath path) (sprintf "Not an include path - '%s'" path)

    let assertIsIncludeDocument (doc:Document<'a>) : DocMonad<'res, unit> = 
        assertM (isIncludeDocument doc) (sprintf "Not an include Document - '%s'" doc.Title)

    let getWorkingPathSuffix (absPath:string) : DocMonad<'res,string> = 
        docMonad { 
            do! assertIsWorkingPath absPath
            let! root = askWorkingDirectory ()
            return rightPathComplement root (FilePath(absPath))
        }
            
    let getWorkingDocPathSuffix (doc:Document<'a>) : DocMonad<'res,string> = 
        getWorkingPathSuffix doc.AbsolutePath


    let getSourcePathSuffix (absPath:string) : DocMonad<'res,string> = 
        docMonad { 
            do! assertIsSourcePath absPath
            let! root = askSourceDirectory ()
            return rightPathComplement root (FilePath(absPath))
        }

    let getSourceDocPathSuffix (doc:Document<'a>) : DocMonad<'res,string> = 
        getSourcePathSuffix doc.AbsolutePath

    //let getIncludePathSuffix (absPath:string) : DocMonad<'res,string> = 
    //    docMonad { 
    //        do! assertIsIncludePath absPath
    //        let! root = askIncludeDirectory ()
    //        return rightPathComplement root (FilePath(absPath))
    //    }

    //let getIncludeDocPathSuffix (doc:Document<'a>) : DocMonad<'res,string> = 
    //    getIncludePathSuffix doc.LocalPath

    //let getPathSuffix (absPath:string) : DocMonad<'res,string> =
    //    getWorkingPathSuffix absPath 
    //        <||> getSourcePathSuffix absPath 
    //        <||> getIncludePathSuffix absPath 

    //let getDocPathSuffix (doc:Document<'a>) : DocMonad<'res,string> =
    //    getWorkingDocPathSuffix doc 
    //        <||> getSourceDocPathSuffix doc
    //        <||> getIncludeDocPathSuffix doc

    /// Create a subdirectory in the Working directory.
    let createWorkingFolder (subfolderName:string) : DocMonad<'res,unit> = 
        askWorkingDirectoryPath () |>> fun cwd ->
        let absFolderName = cwd </> subfolderName
        if Directory.Exists(absFolderName) then
            ()
        else Directory.CreateDirectory(absFolderName) |> ignore
            

    ///// Rewrite the the file name to site it in the working folder.
    ///// If the file is from Source or Include directories generate the name with 
    ///// the respective subfolder path from root.
    ///// Otherwise, generate the file name at the top level of Workgin.
    //let generateWorkingFileName (includeDirectoriesSuffix:bool) (absPath:string) : DocMonad<'res,string> = 
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
    let private copyFileToWorking (sourceAbsPath:string) : DocMonad<'res,Document<'a>> = 
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
    let copyToWorking (doc:Document<'a>) : DocMonad<'res,Document<'a>> = 
        let title = doc.Title
        copyFileToWorking doc.AbsolutePath |>> setTitle title




    /// Rename a folder in the working drectory
    let renameWorkingFolder (oldName:string) (newName:string) : DocMonad<'res,unit> = 
        askWorkingDirectoryPath () >>= fun cwd ->
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
                               (recurseIntoSubDirectories:bool) : DocMonad<'res, bool> = 
        let proc () = 
            askSourceDirectoryPath () |>>  FakeLikePrim.hasFilesMatching pattern recurseIntoSubDirectories
        attemptM <| proc ()



 
    /// Search file matching files in the SourceDirectory.
    /// Uses glob pattern - the only wild cards are '?' and '*'
    /// Returns a list of absolute paths.
    let findAllSourceFilesMatching (pattern:string) 
                                   (recurseIntoSubDirectories:bool) : DocMonad<'res, string list> =
        let proc () = 
            askSourceDirectoryPath () |>> FakeLikePrim.findAllFilesMatching pattern recurseIntoSubDirectories
        attemptM <| proc ()

    /// Search file matching files in the SourceDirectory.
    /// Uses glob pattern - the only wild cards are '?' and '*'
    /// Returns a list of absolute paths.
    let findSomeSourceFilesMatching (pattern:string) 
                                    (recurseIntoSubDirectories:bool) : DocMonad<'res, string list> =
        let proc () = 
            askSourceDirectoryPath () |>> FakeLikePrim.tryFindSomeFilesMatching pattern recurseIntoSubDirectories
        optionFailM "fail - findSomeSourceFilesMatching" <| proc ()

    /// Search file matching files in the SourceDirectory.
    /// Uses glob pattern - the only wild cards are '?' and '*'
    /// Returns a list of absolute paths.
    let tryFindExactlyOneSourceFileMatching (pattern:string) 
                                   (recurseIntoSubDirectories:bool) : DocMonad<'res, string option> =
        let proc () = 
            askSourceDirectoryPath () |>> FakeLikePrim.tryFindExactlyOneFileMatching pattern recurseIntoSubDirectories
        attemptM <| proc ()


    /// Create a subdirectory under the working folder.
    let createWorkingSubdirectory (relPath:string) : DocMonad<'res,unit> = 
        askWorkingDirectoryPath () |>> fun cwd ->
        let path = cwd </> relPath 
        if Directory.Exists(path) then 
            ()
        else Directory.CreateDirectory(path) |> ignore


    /// Run an operation in a subdirectory of current working directory.
    /// Creates the subdirectory if it doesn't exist.
    let localWorkingSubdirectory (subdirectory:string) 
                                 (ma:DocMonad<'res,'a>) : DocMonad<'res,'a> = 
        docMonad {
            let! path = extendWorkingPath subdirectory
            let! _ = createWorkingSubdirectory subdirectory
            return! local (fun env -> {env with WorkingDirectory = DirectoryPath(path)}) ma
        }
            
    /// Run an operation with the Source directory restricted to the
    /// supplied subdirectory.
    let localSourceSubdirectory (subdirectory:string) 
                                (ma:DocMonad<'res,'a>) : DocMonad<'res,'a> = 
        docMonad {
            let! path = extendSourcePath subdirectory
            return! local (fun env -> {env with SourceDirectory = DirectoryPath(path)}) ma
        }



    /// Consider this deprecated...
    let commonSubdirectory (subdirectory:string) 
                           (ma:DocMonad<'res,'a>) : DocMonad<'res,'a> = 
        localWorkingSubdirectory subdirectory <| localSourceSubdirectory subdirectory ma



