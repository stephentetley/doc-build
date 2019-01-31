// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Base


module FakeLikePrim = 

    open System.IO
    open System

    let private traversalOptions (recurseIntoSubDirectories:bool) : SearchOption = 
        if recurseIntoSubDirectories then 
            SearchOption.AllDirectories 
        else 
            SearchOption.TopDirectoryOnly

    /// Uses glob pattern - the only wild cards are '?' and '*'
    let private getFilesMatching (pattern:string)  
                                 (recurseIntoSubDirectories:bool) 
                                 (sourceDirectory:string) : string list =
        DirectoryInfo(sourceDirectory).GetFiles(searchPattern = pattern, searchOption = traversalOptions recurseIntoSubDirectories) 
            |> Array.map (fun (info:FileInfo)  -> info.FullName)
            |> Array.toList

    
    let private tryOneOrMore (input:'a list) : option<'a list> = 
        match input with
        | [] -> None
        | _ -> Some input

    let private tryExactlyOne (input:'a list) : option<'a> = 
        match input with
        | [x] -> Some x
        | _ -> None



    /// Has one or more matches. 
    /// Note - pattern is a simple glob 
    /// (the only wild cards are '?' and '*'), not a regex.
    let hasFilesMatching (pattern:string) 
                         (recurseIntoSubDirectories:bool)
                         (dir:string) : bool = 
        let test = not << List.isEmpty
        getFilesMatching pattern recurseIntoSubDirectories dir |> test

    /// Zero or more matches.
    /// No need for a try variant (empty list is no matches)
    /// Note - pattern is a glob, not a regex.
    let findAllFilesMatching (pattern:string)
                             (recurseIntoSubDirectories:bool) 
                             (dir:string) : string list = 
        getFilesMatching pattern recurseIntoSubDirectories dir

    
    /// One or more matches. 
    /// Note - pattern is a glob, not a regex.
    let tryFindSomeFilesMatching (pattern:string) 
                                 (recurseIntoSubDirectories:bool) 
                                 (dir:string) : option<string list> = 
        getFilesMatching pattern recurseIntoSubDirectories dir |> tryOneOrMore

    /// Exactly one matches.
    /// Note - pattern is a glob, not a regex.
    let tryFindExactlyOneFileMatching (pattern:string) 
                                      (recurseIntoSubDirectories:bool) 
                                      (dir:string) : option<string> = 
        getFilesMatching pattern recurseIntoSubDirectories dir |> tryExactlyOne


    let subdirectoriesWithMatches (pattern:string) (dir:string) : string list = 
        let dirs = System.IO.Directory.GetDirectories(dir) |> Array.toList
        List.filter (hasFilesMatching pattern true) dirs


    let copyFile (target:string) (sourceFile:string) : unit =
        if File.Exists(sourceFile) then 
            let src = FileInfo(sourceFile)
            if Directory.Exists(target) then 
                let target1 = Path.Combine(target,src.Name)
                src.CopyTo(target1) |> ignore
            else
                src.CopyTo(target) |> ignore
        else
            failwithf "copyFile: src not found '%s'" sourceFile

    let fileExists (filePath:string) : bool = 
        File.Exists(filePath) 
    
    let directoryExists (filePath:string) : bool =
        Directory.Exists(filePath)

    /// Recursively delete directory
    let deleteDirectory (dirPath:string) : unit = 
        if Directory.Exists(dirPath) then
            Directory.Delete(path = dirPath, recursive = true)
        else 
            ()



    