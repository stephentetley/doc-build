// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocBuild.Base.Internal


module FilePaths = 

    open System.IO

    type FilePath = string

    /// Is the path a "folder path"?
    /// This does not check for existence, it only checks that 
    /// the path is a folder path (i.e. ends in "/" or "\\")
    let isDirectoryPath (path:string) = 
        // Get the "full path".
        // This normalizes the path to use backslash (Windows style).
        // It horribly adds a prefix to relative files to root them in the users
        // "home folder" but that doesn't matter here.
        let fullname = Path.GetFullPath path
        match Path.GetFileName fullname with
        | null | "" -> true
        | _ -> false

    let isFilePath (path:string) = 
        // Get the "full path".
        // This normalizes the path to use backslash (Windows style).
        // It horribly adds a prefix to relative files to root them in the users
        // "home folder" but that doesn't matter here.
        let fullname = Path.GetFullPath path
        match Path.GetFileName fullname with
        | null | "" -> false
        | _ -> true


    /// Get the "child name".
    /// This is oblivious as to whether the path represents to a file or a folder.
    let getPathName1 (path:string) : string = 
        if isDirectoryPath path then 
            Path.GetDirectoryName path
        else
            Path.GetFileName path


    let private directoryStep (directory:DirectoryInfo) (initialAcc:string list) : string list =
        let rootName = directory.Root.Name
        let rec work (currentDir:DirectoryInfo) (acc:string list) = 
            let folderName = currentDir.Name
            if folderName = rootName then
                rootName :: acc
            else
                work currentDir.Parent (folderName :: acc)
        work directory initialAcc
            

    let pathToSegments (path:string) : option<string list> = 
        if isDirectoryPath path then 
            directoryStep (new DirectoryInfo(path)) [] |> Some
        elif isFilePath path then 
            let fileInfo = new FileInfo(path)
            directoryStep fileInfo.Directory [fileInfo.Name] |> Some
        else  None


    let segmentsToPath (segments:string list) : string = 
        Path.Combine(paths = Array.ofList segments)

    
    // Maybe the functions (and their names) are not a good enough API.

    let commonPathPrefix (path1:string) (path2:string) : string = 
        let segments1 = pathToSegments path1 |> Option.defaultValue [] 
        let segments2 = pathToSegments path2 |> Option.defaultValue [] 
        let rec work xs ys (acc:string list) = 
            match xs, ys with
            | (s :: ss), (t :: ts) -> 
                if s = t then 
                    work ss ts (s :: acc)
                else
                    List.rev acc
            | _, _ -> List.rev acc
        work segments1 segments2 [] |> segmentsToPath



    let rightPathComplement (ancestor:string) (child:string) : string = 
        let segments1 = pathToSegments ancestor |> Option.defaultValue [] 
        let segments2 = pathToSegments child    |> Option.defaultValue [] 
        let rec work xs ys  = 
            match xs, ys with
            | (s :: ss), (t :: ts) -> 
                if s = t then 
                    work ss ts
                else ys
            | _, _ -> ys
        work segments1 segments2 |> segmentsToPath


    let rootIsPrefix (root:string) (child:string) : bool = 
        let segments1 = pathToSegments root     |> Option.defaultValue [] 
        let segments2 = pathToSegments child    |> Option.defaultValue [] 
        let rec work xs ys  = 
            match xs, ys with
            | (s :: ss), (t :: ts) -> 
                if s = t then 
                    work ss ts
                else false
            | [], _ -> true
            | _, _ -> false
        work segments1 segments2


