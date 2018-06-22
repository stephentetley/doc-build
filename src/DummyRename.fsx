// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

// FAKE dependencies are getting onorous...
#I @"..\packages\FAKE.5.0.0-rc016.225\tools"
#r @"FakeLib.dll"
#I @"..\packages\Fake.Core.Globbing.5.0.0-beta021\lib\net46"
#r @"Fake.Core.Globbing.dll"
#I @"..\packages\Fake.IO.FileSystem.5.0.0-rc017.237\lib\net46"
#r @"Fake.IO.FileSystem.dll"
#I @"..\packages\Fake.Core.Trace.5.0.0-rc017.237\lib\net46"
#r @"Fake.Core.Trace.dll"
#I @"..\packages\Fake.Core.Process.5.0.0-rc017.237\lib\net46"
#r @"Fake.Core.Process.dll"

#load @"DocMake\Base\Common.fs"
#load @"DocMake\Base\FakeLike.fs"
#load @"DocMake\Base\CopyRename.fs"
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



