// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Base

[<AutoOpen>]
module FindFiles = 

    open System.IO


    let tryExactlyOne (source : 'a list) : 'a option = 
        match source with
        | [a] -> Some a
        | _ -> None

    let tryAtLeastOne (source : 'a list) : option<'a list> = 
        match source with
        | [] -> None
        | xs -> Some xs


    let private searchOpts (recurseIntoSubdirectories:bool) : SearchOption = 
        if recurseIntoSubdirectories then 
            SearchOption.AllDirectories 
        else 
            SearchOption.TopDirectoryOnly

    /// Uses glob pattern - the only wild cards are '?' and '*'
    let private getFilesMatching (pattern : string)  
                                 (recurseIntoSubdirectories : bool) 
                                 (sourceDirectory : string) : string list =
        let options = searchOpts recurseIntoSubdirectories
        DirectoryInfo(sourceDirectory).GetFiles( searchPattern = pattern
                                               , searchOption = options ) 
            |> Array.map (fun (info:FileInfo)  -> info.FullName)
            |> Array.toList

    /// Uses glob pattern - the only wild cards are '?' and '*'
    let private hasFilesMatching (pattern : string)  
                                 (recurseIntoSubdirectories : bool) 
                                 (sourceDirectory:string) : bool =
        getFilesMatching pattern recurseIntoSubdirectories sourceDirectory 
            |> (not << List.isEmpty)


    /// Has one or more matches. 
    /// Note - pattern is a simple glob 
    /// (the only wild cards are '?' and '*'), not a regex.
    let hasSourceFilesMatching (pattern:string) 
                               (recurseIntoSubDirectories:bool) : DocMonad<bool, 'userRes> = 
        docMonad { 
            let! source = askSourceDirectory ()
            return! liftOperation "hasFilesMatching error" 
                            (fun () -> hasFilesMatching pattern recurseIntoSubDirectories source)
        }

    /// Search file matching files in the SourceDirectory.
    /// Uses glob pattern - the only wild cards are '?' and '*'
    /// Returns a list of absolute paths.
    let findSourceFilesMatching (pattern:string) 
                                (recurseIntoSubDirectories:bool) : DocMonad<string list, 'userRes> =
        docMonad { 
            let! source = askSourceDirectory ()
            return! liftOperation "no matches" 
                            (fun _ -> getFilesMatching pattern recurseIntoSubDirectories source)
        }