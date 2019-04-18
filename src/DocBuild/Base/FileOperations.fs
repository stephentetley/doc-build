﻿// Copyright (c) Stephen Tetley 2019
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
        askWorkingDirectory () |>> fun dir -> rootIsPrefix dir (FilePath(absPath))


    let isWorkingDocument (doc:Document<'a>) : DocMonad<'res,bool> = 
        isWorkingPath doc.LocalPath

    let isSourcePath (absPath:string) : DocMonad<'res,bool> = 
        askSourceDirectory () |>> fun dir -> rootIsPrefix dir (FilePath(absPath))
        

    let isSourceDocument (doc:Document<'a>) : DocMonad<'res,bool> = 
        isSourcePath doc.LocalPath


    let isIncludePath (absPath:string) : DocMonad<'res,bool> = 
        askIncludeDirectory () |>> fun dir -> rootIsPrefix dir (FilePath(absPath))
        
    
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
            let! root = askWorkingDirectory ()
            return rightPathComplement root (FilePath(absPath))
        }
            
    let getWorkingDocPathSuffix (doc:Document<'a>) : DocMonad<'res,string> = 
        getWorkingPathSuffix doc.LocalPath


    let getSourcePathSuffix (absPath:string) : DocMonad<'res,string> = 
        docMonad { 
            do! assertIsSourcePath absPath
            let! root = askSourceDirectory ()
            return rightPathComplement root (FilePath(absPath))
        }

    let getSourceDocPathSuffix (doc:Document<'a>) : DocMonad<'res,string> = 
        getSourcePathSuffix doc.LocalPath

    let getIncludePathSuffix (absPath:string) : DocMonad<'res,string> = 
        docMonad { 
            do! assertIsIncludePath absPath
            let! root = askIncludeDirectory ()
            return rightPathComplement root (FilePath(absPath))
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
    let generateWorkingFileName (includeDirectoriesSuffix:bool) (absPath:string) : DocMonad<'res,string> = 
        docMonad { 
            let! suffix = 
                if includeDirectoriesSuffix then 
                    getPathSuffix absPath <||> mreturn (FileInfo(absPath).Name)
                else mreturn (FileInfo(absPath).Name)
            return! extendWorkingPath suffix
        }

    /// Copy a file to working, returning the copy as a Document.
    /// If the file is from Source or Include directories copy with 
    /// the respective subfolder path from root.
    let copyFileToWorking (includeDirectoriesSuffix:bool) (sourceAbsPath:string) : DocMonad<'res,Document<'a>> = 
        docMonad { 
            let! target = generateWorkingFileName includeDirectoriesSuffix sourceAbsPath
            if File.Exists(target) then 
                File.Delete(target) 
            else ()
            File.Copy( sourceFileName = sourceAbsPath, destFileName = target )
            return Document(target)
        }

    /// Copy a doc to working.
    /// If the file is from Source or Include copy with the respective
    /// subfolder path from root.
    let copyToWorking (includeDirectoriesSuffix:bool) (doc:Document<'a>) : DocMonad<'res,Document<'a>> = 
        copyFileToWorking includeDirectoriesSuffix doc.LocalPath




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

    /// Run an operation with the Include directory restricted to the
    /// supplied sub-directory.
    let localIncludeSubdirectory (subdirectory:string) 
                                 (ma:DocMonad<'res,'a>) : DocMonad<'res,'a> = 
        docMonad {
            let! path = extendIncludePath subdirectory
            return! local (fun env -> {env with IncludeDirectory = DirectoryPath(path)}) ma
        }

    /// Consider this deprecated...
    let commonSubdirectory (subdirectory:string) 
                           (ma:DocMonad<'res,'a>) : DocMonad<'res,'a> = 
        localWorkingSubdirectory subdirectory <| localSourceSubdirectory subdirectory ma


    let copyFileToWorkingSubdirectory (subdirectory:string) (srcAbsPath:string) : DocMonad<'res,Document<'a>> = 
        localWorkingSubdirectory subdirectory 
            <| copyFileToWorking false srcAbsPath


    let private getSourceChildren () : DocMonad<'res,string list> = 
        let failMessage = fun _ -> "getSourceChildren directory error"
        docMonad { 
            let! source = askSourceDirectoryPath ()
            let! (kids: IO.DirectoryInfo[]) = 
                liftAction failMessage <| fun _ -> System.IO.DirectoryInfo(source).GetDirectories() 
            return (kids |> Array.map (fun info -> info.Name) |> Array.toList)
        }

    type SkeletonStepBeginMessage = int -> int -> string -> string

    type SkeletonStepFailMessage = string -> string

    type SkeletonOptions = 
        { GenStepBeginMessage: SkeletonStepBeginMessage option
          GenStepFailMessage: SkeletonStepFailMessage option
          DebugSelectSample: (string list -> string list) option
          ContinueOnFail: bool
        }

    let defaultSkeletonOptions:SkeletonOptions = 
        { GenStepBeginMessage = Some 
            <| fun ix count childFolderName -> 
                    sprintf "%i of %i: %s" ix count childFolderName
          GenStepFailMessage = Some <| sprintf "%s failed"
          DebugSelectSample = None
          ContinueOnFail = false
        }

    let private runSkeleton (skeletonOpts:SkeletonOptions) 
                            (strategy: string -> DocMonad<'res,unit> -> DocMonad<'res,unit>)
                            (process1: DocMonad<'res,'a>) : DocMonad<'res, unit> =  
        let processZ: DocMonad<'res, unit> = process1 |>> fun _ -> ()
        let filterChildDirectories = 
            match skeletonOpts.DebugSelectSample with
            | None -> id
            | Some fn -> fn
        let logStepBegin (ix:int) (count:int) = 
            match skeletonOpts.GenStepBeginMessage with
            | None -> mreturn ()
            | Some genMessage -> 
                docMonad { 
                   let! kid = askSourceDirectoryName ()
                   let message = genMessage ix count kid 
                   do (printfn "%s" message)
                   do! tellLine message
                   return ()
                }
        let logStepFail () = 
            match skeletonOpts.GenStepFailMessage with
            | None -> mreturn ()
            | Some genMessage -> 
                docMonad { 
                    let! kid = askSourceDirectoryName ()
                    let message = genMessage kid
                    do (printfn "%s" message)
                    do! tellLine message
                    return ()
                }
        let proceedM (proc:DocMonad<'res,unit>) : DocMonad<'res, unit> = 
            docMonad { 
                match! (optionalM proc) with
                | None -> 
                    if skeletonOpts.ContinueOnFail then 
                        return ()
                    else 
                        logStepFail () .>> throwError "Build step failed" |> ignore
                | Some _ -> return ()
                }
        let processChildDirectory (ix:int) (count:int) : DocMonad<'res, unit> = 
            docMonad { 
                do! logStepBegin ix count
                return! (proceedM processZ)
            }

        getSourceChildren () >>= fun srcDirs -> 
        let sources = filterChildDirectories srcDirs
        let count = List.length sources
        foriMz sources 
               (fun ix dir -> strategy dir (processChildDirectory (ix + 1) count))


    /// Processing skeleton.
    /// For every child source folder (one level down) run the
    /// processing function on 'within' that folder. 
    /// Generate the results in a child folder of the same name under
    /// the working folder.
    let dtodSourceChildren (skeletonOpts:SkeletonOptions) 
                           (process1: DocMonad<'res,'a>) : DocMonad<'res, unit> = 
        let strategy = fun childDirectory action -> 
                localSourceSubdirectory childDirectory 
                                        (localWorkingSubdirectory childDirectory action)
        runSkeleton skeletonOpts strategy process1

    /// Processing skeleton.
    /// For every child source folder (one level down) run the
    /// processing function on 'within' that folder. 
    /// Generate the results in the top level working folder.
    let dto1SourceChildren (skeletonOpts:SkeletonOptions) 
                           (process1: DocMonad<'res,'a>) : DocMonad<'res, unit> = 
        let strategy = fun childDirectory action -> 
                localSourceSubdirectory childDirectory action
        runSkeleton skeletonOpts strategy process1