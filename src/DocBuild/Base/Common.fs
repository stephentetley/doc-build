// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Base

[<AutoOpen>]
module Common = 

    open System
    open System.IO

    open DocBuild.Base.Internal

    type PageOrientation = 
        | Portrait 
        | Landscape



    type PrintQuality = 
        | Screen
        | Print


    // ************************************************************************
    // Find and replace


    type SearchList = (string * string) list


    type ErrMsg = string

    type BuildResult<'a> = Result<'a,ErrMsg>


    let doubleQuote (source:string) : string = 
        sprintf "\"%s\"" source



    // ************************************************************************
    // File path concat


    /// This is System.IO.Path.Combine.
    /// "directory" </> "/file.ext" == "/file.ext"
    let ( </> ) (path1:string) (path2:string) : string = 
        Path.Combine(path1, path2)

    /// Get the "leaf name" of a path omitting any directory prefix.
    /// This is oblivious as to whether the path represents to a file or a folder.
    let fileObjectName (path:string) : string = 
        FilePaths.getPathName1 path


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








    // ************************************************************************
    // File name helpers


    let safeName (input:string) : string = 
        let parens = ['('; ')'; '['; ']'; '{'; '}']
        let bads = ['\\'; '/'; ':'; '?'; '*'] 
        let white = ['\n'; '\t']
        let ans1 = List.fold (fun (s:string) (c:char) -> s.Replace(c.ToString(), "")) input parens
        let ans2 = List.fold (fun (s:string) (c:char) -> s.Replace(c,'_')) ans1 bads
        let ans3 = List.fold (fun (s:string) (c:char) -> s.Replace(c,'_')) ans2 white
        ans3.Trim() 


    /// Suffix a file name _before_ the extension.
    ///
    /// e.g suffixFileName "TEMP"  "sunset.jpg" ==> "sunset.TEMP.jpg"
    let suffixFileName (suffix:string)  (filePath:string) : string = 
        let root = System.IO.Path.GetDirectoryName filePath
        let justfile = System.IO.Path.GetFileNameWithoutExtension filePath
        let ext  = System.IO.Path.GetExtension filePath
        let newfile = sprintf "%s.%s%s" justfile suffix ext
        Path.Combine(root, newfile)
