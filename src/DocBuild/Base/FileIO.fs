// Copyright (c) Stephen Tetley 2019
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


    let askWorkingDirectory () : DocMonad<'res,string> = 
        asks (fun env -> env.WorkingDirectory)

    let askSourceDirectory () : DocMonad<'res,string> = 
        asks (fun env -> env.SourceDirectory)
        
    let askIncludeDirectory () : DocMonad<'res,string> = 
        asks (fun env -> env.IncludeDirectory)

    let private askFile (getTopLevel:unit -> DocMonad<'res,string>)
                        (fileName:string) : DocMonad<'res,string> = 
        docMonad { 
            let! cwd = getTopLevel ()
            let path = cwd </> fileName
            return path
        }

    /// Return the full path of a filename local to the working directory.
    /// Does not validate if the file exists
    let askWorkingFile (fileName:string) : DocMonad<'res,string> = 
        askFile askWorkingDirectory fileName


    /// Return the full path of a filename local to the Source directory.
    /// Does not validate if the file exists
    let askSourceFile (fileName:string) : DocMonad<'res,string> = 
        askFile askSourceDirectory fileName


    /// Return the full path of a filename local to the Include directory.
    /// Does not validate if the file exists
    let askIncludeFile (fileName:string) : DocMonad<'res,string> = 
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
            let path = cwd </> subDirectory
            do! attempt (create1 path)
        }

    /// Run an operation in a subdirectory of current working directory.
    /// Create the directory if it doesn't exist.
    let localSubDirectory (subDirectory:string) 
                          (ma:DocMonad<'res,'a>) : DocMonad<'res,'a> = 
        docMonad {
            let! cwd = askWorkingDirectory ()
            let path = cwd </> subDirectory
            do! createWorkingSubDirectory path
            let! ans = local (fun env -> {env with WorkingDirectory = path}) ma
            return ans
        }

    /// Run an operation with the Source directory restricted to the
    /// supplied sub-directory.
    let childSourceDirectory (subDirectory:string) 
                             (ma:DocMonad<'res,'a>) : DocMonad<'res,'a> = 
        docMonad {
            let! srcDir = askSourceDirectory ()
            let path = srcDir </> subDirectory
            let! ans = local (fun env -> {env with SourceDirectory = path}) ma
            return ans
        }

    /// This is a bit too primitive, ideally it would work on Documents.
    let copyToWorking (doc:Document<'a>) : DocMonad<'res,Document<'a>> = 
        if File.Exists(doc.Path) then 
            docMonad { 
                let justFile = Path.GetFileName(doc.Path)
                let! cwd = askWorkingDirectory ()
                let target = cwd </> justFile
                do File.Copy( sourceFileName = doc.Path
                            , destFileName = target )
                return Document(target)
            }
        else
            throwError 
                <| sprintf "copyToWorking: sourceFile not found '%s'" doc.Path

    let localFile (fileName:string) : DocMonad<'res,string> = 
        docMonad { 
            let! cwd = askWorkingDirectory ()
            return (cwd </> fileName)
        }            

    // ************************************************************************
    // Source files
            
    /// Search file matching files in the SourceDirectory.
    /// Uses glob pattern - the only wild cards are '?' and '*'
    let getSourceFilesMatching (pattern:string) : DocMonad<'res, string list> =
        askSourceDirectory () |>> (fun src -> FakeLike.getFilesMatching src pattern)
            