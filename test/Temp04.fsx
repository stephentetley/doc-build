// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

#r "netstandard"

open System.IO
open System


let cwd = @"D:\coding\fsharp\doc-build\data"
let path1 = @"D:\coding\fsharp\doc-build\..\doc-build\data\temp1.pdf"

// GetFullPath does sufficient normalization.
// But we should use Uri instead.
let test01 () =
    Path.GetFullPath path1

let test02 () =
    let uPath1 = new Uri(path1)
    let uCwd = new Uri(cwd)
    printfn "%O" <| uCwd.IsBaseOf(uPath1)


// Don't bother with Uris for relative file paths.
let test03 () = 
    let u1 = new System.Uri(uriString= @"local\file.txt", uriKind=UriKind.Relative)
    u1






