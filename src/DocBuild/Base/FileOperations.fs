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


    let getOutputPath (relativeFileName:string) : DocMonad<'res,string> = 
        askWorkingDirectory () |>> fun cwd -> (cwd.LocalPath </> relativeFileName)


    let askWorkingDirectoryPath () : DocMonad<'res,string> = 
        askWorkingDirectory () |>> fun uri -> uri.LocalPath

    let askSourceDirectoryPath () : DocMonad<'res,string> = 
        askSourceDirectory () |>> fun uri -> uri.LocalPath

    let askIncludeDirectoryPath () : DocMonad<'res,string> = 
        askIncludeDirectory () |>> fun uri -> uri.LocalPath

    let extendWorkingPath (relPath:string) : DocMonad<'res,string> = 
        askWorkingDirectoryPath () |>> fun root -> root </> relPath

    let extendSourcePath (relPath:string) : DocMonad<'res,string> = 
        askSourceDirectoryPath () |>> fun root -> root </> relPath
    
    let extendIncludePath (relPath:string) : DocMonad<'res,string> = 
        askIncludeDirectoryPath () |>> fun root -> root </> relPath

    let isWorkingPath (absPath:string) : DocMonad<'res,bool> = 
        askWorkingDirectory () |>> fun dir -> dir.IsBaseOf(new Uri(absPath))


    let isWorkingDocument (doc:Document<'a>) : DocMonad<'res,bool> = 
        isWorkingPath doc.LocalPath

    let isSourcePath (absPath:string) : DocMonad<'res,bool> = 
        askSourceDirectory () |>> fun dir -> dir.IsBaseOf(new Uri(absPath))
        

    let isSourceDocument (doc:Document<'a>) : DocMonad<'res,bool> = 
        isSourcePath doc.LocalPath


    let isIncludePath (absPath:string) : DocMonad<'res,bool> = 
        askIncludeDirectory () |>> fun dir -> dir.IsBaseOf(new Uri(absPath))
        
    
    let isIncludeDocument (doc:Document<'a>) : DocMonad<'res,bool> = 
        isIncludePath doc.LocalPath

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
            let! dir = askWorkingDirectory ()
            let rel = dir.MakeRelativeUri(new Uri(absPath))
            return rel.ToString()
        }
            
    let getWorkingDocPathSuffix (doc:Document<'a>) : DocMonad<'res,string> = 
        getWorkingPathSuffix doc.LocalPath


    let getSourcePathSuffix (absPath:string) : DocMonad<'res,string> = 
        docMonad { 
            do! assertIsSourcePath absPath
            let! dir = askSourceDirectory ()
            let rel = dir.MakeRelativeUri(new Uri(absPath))
            return rel.ToString()
        }

    let getSourceDocPathSuffix (doc:Document<'a>) : DocMonad<'res,string> = 
        getSourcePathSuffix doc.LocalPath

    let getIncludePathSuffix (absPath:string) : DocMonad<'res,string> = 
        docMonad { 
            do! assertIsIncludePath absPath
            let! dir = askIncludeDirectory ()
            let rel = dir.MakeRelativeUri(new Uri(absPath))
            return rel.ToString()
        }

    let getIncludeDocPathSuffix (doc:Document<'a>) : DocMonad<'res,string> = 
        getIncludePathSuffix doc.LocalPath

    let getPathSuffix (absPath:string) : DocMonad<'res,string> =
        getWorkingPathSuffix absPath 
            <||> getSourcePathSuffix absPath 
            <||> getIncludePathSuffix absPath 

    let getDocPathSuffix (doc:Document<'a>) : DocMonad<'res,string> =
        getWorkingDocPathSuffix doc 
            <||> getSourceDocPathSuffix doc
            <||> getIncludeDocPathSuffix doc

    /// Create a subdirectory in the Working directory.
    let createWorkingFolder (subfolderName:string) : DocMonad<'res,unit> = 
        askWorkingDirectoryPath () |>> fun cwd ->
        let absFolderName = cwd </> subfolderName
        if Directory.Exists(absFolderName) then
            ()
        else Directory.CreateDirectory(absFolderName) |> ignore
            

    /// Rewrite the the file name to site it in the working folder.
    /// If the file is from Source or Include directories generate the name with 
    /// the respective subfolder path from root.
    /// Otherwise, generate the file name at the top level of Workgin.
    let generateWorkingFileName (absPath:string) : DocMonad<'res,string> = 
        docMonad { 
            let! suffix = getPathSuffix absPath <||> dreturn (FileInfo(absPath).Name)
            return! extendWorkingPath suffix
        }

    /// Copy a file to working, returning the copy as a Document.
    /// If the file is from Source or Include directories copy with 
    /// the respective subfolder path from root.
    let copyFileToWorking (absPath:string) : DocMonad<'res,Document<'a>> = 
        docMonad { 
            let! target = generateWorkingFileName absPath
            if File.Exists(target) then 
                File.Delete(target) 
            else ()
            File.Copy( sourceFileName = absPath, destFileName = target )
            return Document(target)
        }

    /// Copy a doc to working.
    /// If the file is from Source or Include copy with the respective
    /// subfolder path from root.
    let copyToWorking (doc:Document<'a>) : DocMonad<'res,Document<'a>> = 
        copyFileToWorking doc.LocalPath


    /// Rename a folder in the working drectory
    let renameWorkingFolder (oldName:string) (newName:string) : DocMonad<'res,unit> = 
        askWorkingDirectoryPath () >>= fun cwd ->
        let oldPath = cwd </> oldName
        if Directory.Exists(oldPath) then
            let newPath = cwd </> newName
            Directory.Move(oldPath, newPath)
            dreturn ()
        else
            throwError (sprintf "renameWorkingFolder - folder does not exist '%s'" oldPath)

    /// Has one or more matches. 
    /// Note - pattern is a simple glob 
    /// (the only wild cards are '?' and '*'), not a regex.
    let hasSourceFilesMatching (pattern:string) 
                               (recurseIntoSubDirectories:bool) : DocMonad<'res, bool> = 
        askSourceDirectoryPath () |>>  FakeLikePrim.hasFilesMatching pattern recurseIntoSubDirectories




 
    /// Search file matching files in the SourceDirectory.
    /// Uses glob pattern - the only wild cards are '?' and '*'
    /// Returns a list of absolute paths.
    let findAllSourceFilesMatching (pattern:string) 
                                   (recurseIntoSubDirectories:bool) : DocMonad<'res, string list> =
        askSourceDirectoryPath () |>> FakeLikePrim.findAllFilesMatching pattern recurseIntoSubDirectories


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
            return! local (fun env -> {env with WorkingDirectory = new Uri(path)}) ma
        }
            
    /// Run an operation with the Source directory restricted to the
    /// supplied subdirectory.
    let localSourceSubdirectory (subdirectory:string) 
                                (ma:DocMonad<'res,'a>) : DocMonad<'res,'a> = 
        docMonad {
            let! path = extendSourcePath subdirectory
            return! local (fun env -> {env with SourceDirectory = new Uri(path)}) ma
        }

    /// Run an operation with the Include directory restricted to the
    /// supplied sub-directory.
    let localIncludeSubdirectory (subdirectory:string) 
                                 (ma:DocMonad<'res,'a>) : DocMonad<'res,'a> = 
        docMonad {
            let! path = extendIncludePath subdirectory
            return! local (fun env -> {env with IncludeDirectory = new Uri(path)}) ma
        }

  