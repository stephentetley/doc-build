// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


#load @"..\src\DocMake\Base\Common.fs"
#load @"..\src\DocMake\Base\FakeLike.fs"
#load @"..\src\DocMake\Base\CopyRename.fs"
open DocMake.Base.FakeLike
open DocMake.Base.CopyRename

let test01 () = 
    let fmt = Printf.StringFormat<int->string> "dummy-file_%03i.txt"
    sprintf fmt 1

let test02 () =
    let files = findAllMatchingFiles @"G:\work\photos1\TestFolder" "\.jpg$"
    List.iter (fun a -> printfn "%s" a) files

let test03 () =
    let fmt1 = Printf.StringFormat<int->string> "Photo%04i.jpg" 
    let srcDir = @"G:\work\photos1\TestFolder"
    let destDir = @"G:\work\photos1\TestFolder\output"
    multiCopyGlobRename (srcDir, "DSCF*.jpg") (destDir, sprintf fmt1)



