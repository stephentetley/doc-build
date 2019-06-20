// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Base

[<AutoOpen>]
module FileOperations = 

    open System.IO
    open System

    open DocBuild.Base.Internal
    open DocBuild.Base

    /// TODO - need to unify what file ops we actaully need


    let extendWorkingPath (relativePath:string) : DocMonad<string, 'userRes> = 
        askWorkingDirectory () |>> fun root -> root </> relativePath

    let extendSourcePath (relativePath:string) : DocMonad<string, 'userRes> = 
        askSourceDirectory () |>> fun root -> root </> relativePath
    



    /// Create a subdirectory in the Working directory.
    let createWorkingSubdirectory (subdirectory : string) : DocMonad<unit, 'userRes> = 
        askWorkingDirectory () |>> fun cwd ->
        let absFolderName = cwd </> subdirectory
        if Directory.Exists(absFolderName) then
            ()
        else Directory.CreateDirectory(absFolderName) |> ignore
            



    /// Copy a file to working, returning the copy as a Document.
    /// If the file is from Source or Include directories copy with 
    /// the respective subfolder path from root.
    let private copyFileToWorking (sourceAbsPath:string) : DocMonad<Document<'a>, 'userRes> = 
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
    let copyDocumentToWorking (doc:Document<'a>) : DocMonad<Document<'a>, 'userRes> = 
        let title = doc.Title
        copyFileToWorking doc.AbsolutePath |>> setTitle title




    /// Rename a folder in the working drectory
    let renameWorkingSubdirectory (oldName:string) (newName:string) : DocMonad<unit, 'userRes> = 
        askWorkingDirectory () >>= fun cwd ->
        let oldPath = cwd </> oldName
        if Directory.Exists(oldPath) then
            let newPath = cwd </> newName
            Directory.Move(oldPath, newPath)
            mreturn ()
        else
            docError (sprintf "renameWorkingFolder - folder does not exist '%s'" oldPath)





    /// Run an operation in a subdirectory of current working directory.
    /// Creates the subdirectory if it doesn't exist.
    let localWorkingSubdirectory (subdirectory:string) 
                                 (ma:DocMonad<'a, 'userRes>) : DocMonad<'a, 'userRes> = 
        docMonad {
            let! path = extendWorkingPath subdirectory
            let! _ = createWorkingSubdirectory subdirectory
            return! local (fun env -> {env with WorkingDirectory = path}) ma
        }
            


    /// Run an operation with the Source directory restricted to the
    /// supplied subdirectory.
    let localSourceSubdirectory (subdirectory:string) 
                                (ma:DocMonad<'a, 'userRes>) : DocMonad<'a, 'userRes> = 
        docMonad {
            let! path = extendSourcePath subdirectory
            return! local (fun env -> {env with SourceDirectory = path}) ma
        }






