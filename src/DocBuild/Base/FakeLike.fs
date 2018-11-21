// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


namespace DocBuild.Base.FakeLike

[<AutoOpen>]
module FakeLike = 

    open System.IO


    /// Note if the second path is prefixed by '\\'
    /// "directory" </> "/file.ext" == "/file.ext"
    let (</>) (path1:string) (path2:string) = 
        Path.Combine(path1, path2)

    /// This replaces (!!) for simple cases - the only wild cards are '?' and '*'
    let getFilesMatching (sourceDirectory:string) (pattern:string) : string list =
        DirectoryInfo(sourceDirectory).GetFiles(searchPattern = pattern) 
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



    // Has one or more matches. 
    // Note - pattern is a simple glob (the only wild cards are '?' and '*'), not a regex.
    let hasMatchingFiles (pattern:string) (dir:string) : bool = 
        let test = not << List.isEmpty
        getFilesMatching dir pattern |> test

    // Zero or more matches.
    // No need for a try variant (empty list is no matches)
    // Note - pattern is a glob, not a regex.
    let findAllMatchingFiles (pattern:string) (dir:string) : string list = 
        getFilesMatching dir pattern

    
    // One or more matches. 
    // Note - pattern is a glob, not a regex.
    let tryFindSomeMatchingFiles (pattern:string) (dir:string) : option<string list> = 
        getFilesMatching dir pattern |> tryOneOrMore

    // Exactly one matches.
    // Note - pattern is a glob, not a regex.
    let tryFindExactlyOneMatchingFile (pattern:string) (dir:string) : option<string> = 
        getFilesMatching dir pattern |> tryExactlyOne


    let subdirectoriesWithMatches (pattern:string) (dir:string) : string list = 
        let dirs = System.IO.Directory.GetDirectories(dir) |> Array.toList
        List.filter (hasMatchingFiles pattern) dirs

    
    let assertMandatory (failMsg:string) : unit = failwithf "FAIL: Mandatory: %s" failMsg

    let assertOptional  (warnMsg:string) : unit = printfn "WARN: Optional: %s" warnMsg

    let executeProcess (toolPath:string) (command:string) : int = 
        try
            let procInfo = new System.Diagnostics.ProcessStartInfo ()
            procInfo.FileName <- toolPath
            procInfo.Arguments <- command
            procInfo.CreateNoWindow <- true
            let proc = new System.Diagnostics.Process()
            proc.StartInfo <- procInfo
            proc.Start() |> ignore
            proc.WaitForExit () 
            proc.ExitCode
        with
        | ex -> failwith (sprintf "executeProcess: \n%s" ex.Message)


    let copyFile (target:string) (sourceFile:string) : unit =
        if File.Exists(sourceFile) then 
            let src = FileInfo(sourceFile)
            if Directory.Exists(target) then 
                let target1 = target </> src.Name
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



    