// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Base

[<AutoOpen>]
module Common = 

    open System
    open System.IO

    type PageOrientation = 
        | OrientationPortrait 
        | OrientationLandscape



    type PrintQuality = 
        | PqScreen
        | PqPrint


    type ErrMsg = string

    type BuildResult<'a> = Result<'a,ErrMsg>


    let doubleQuote (source:string) : string = 
        sprintf "\"%s\"" source

    // ************************************************************************
    // File path concat


    /// This is just an alias for System.IO.Path.Combine.
    /// Note if the second path is prefixed by '\\'
    /// "directory" </> "/file.ext" == "/file.ext"
    let ( </> ) (path1:string) (path2:string) : string = 
        Path.Combine(path1, path2)

    let ( <//> ) (uri:Uri) (path2:string) : Uri = 
        new Uri (Path.Combine(uri.LocalPath, path2))

    let folderUri (path:string) = 
        let path1 = DirectoryInfo(sprintf "%s%c" path Path.DirectorySeparatorChar).FullName
        new Uri (path1)

    // ************************************************************************
    // Find and replace


    type SearchList = (string * string) list


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
