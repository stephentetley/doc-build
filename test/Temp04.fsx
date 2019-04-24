// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

#r "netstandard"

open System.IO
open System

#load "..\src\DocBuild\Base\FilePaths.fs"
open DocBuild.Base


/// Note cwd is not a directory
let cwd = @"D:\coding\fsharp\doc-build\data"

let cwd1 = @"D:\coding\fsharp\doc-build\data\"
let path1 = @"D:\coding\fsharp\doc-build\..\doc-build\data\temp1.pdf"






// GetFullPath does sufficient normalization.
// But we should use DirectoryPath / FilePath for calculation
let test01 () =
    Path.GetFullPath path1


let test02 () = 
    System.IO.Path.GetDirectoryName(@"folder1/temp.txt") |> printfn "%s"
    System.IO.Path.GetDirectoryName(@"folder1\temp.txt") |> printfn "%s"
    // GetFullPath fixes the name to "user-temp" horrible...
    System.IO.Path.GetFullPath(@"folder1\temp.txt") |> printfn "%s"
    System.IO.Path.Combine([| @"folder1/temp.txt" |]) |> printfn "%s"

let test03 () =
    Path.IsPathRooted cwd |> printfn "%A"

let paradox () = 
    let fullname = Path.GetFullPath "temp.txt"
    printfn "fullname: %s" fullname
    let dirname = Path.GetDirectoryName fullname
    printfn "dirname: %s" dirname
    isDirectoryPath dirname

let test04 () = 
    // Not relevant to handling file paths...
    Path.PathSeparator


    





//let test02 () =
//    let pPath1 = FilePath(path1)
//    let pCwd = DirectoryPath(cwd)
//    printfn "%O" <| rootIsPrefix pCwd pPath1


//// Don't bother with Uris for relative file paths.
//let test03 () = 
//    let u1 = new System.Uri(uriString= @"local\file.txt", uriKind=UriKind.Relative)
//    u1

///// Cannot read .AbsolutePath of an Uri is not right for 
///// the File IO API, we should be using .LocaPath.
//let test04 () = 
//    let pathWithSpaces = @"G:\work\working\FOLDER WITH SPACES\FILE WITH SPACES.txt"
//    let uri = new Uri(pathWithSpaces)
//    printfn "%s" uri.LocalPath
//    System.IO.File.ReadAllText(uri.LocalPath)





///// Using uri is worse than path! 
///// It provides a 'right complement' operation but adds complexity
///// (escaped spaces etc.) that are making the code error prone
///// Solution: write commonPrefix and rightComplement for file paths.

//let zz01 () = 
//    (FilePath @"Z:\oresenna\fsharp\doc-build\..\doc-build\data").LocalPath

//let zz02 () = 
//    (FilePath @"..\doc-build\data")

//let zz03 () = 
//    Path.IsPathRooted @"coding\fsharp\file1.fs"

//let zz03b () = 
//    let relPath = @"coding\fsharp\file1.fs"
//    let fileInfo = FileInfo(relPath)
//    fileInfo.Directory.Root

//let zz04 () = 
//    let path = @"d:\coding\fsharp\file1.fs"
//    FilePath(path).LocalPath


//let zz04b () = 
//    let path = @"d:\coding\fsharp"
//    DirectoryPath(path).LocalPath

//let zz05 () = 
//    commonPathPrefix (DirectoryPath @"d:\coding\fsharp") (FilePath @"d:\coding\fsharp\file1.fs")

//let zz06 () = 
//    rightPathComplement (DirectoryPath @"d:\coding\") (FilePath @"d:\coding\fsharp\file1.fs")
