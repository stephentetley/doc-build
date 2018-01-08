// FAKE is local to the project file
#I @"..\packages\FAKE.5.0.0-beta005\tools"
#r @"..\packages\FAKE.5.0.0-beta005\tools\FakeLib.dll"

#load @"DocMake\Base\Common.fs"
#load @"DocMake\Tasks\UniformRename.fs"
open DocMake.Tasks.UniformRename

let test01 () = 
    let fmt = Printf.StringFormat<int->string> "dummy-file_%03i.txt"
    sprintf fmt 1

let test02 () =
    let files = matchFiles @"G:\work\photos1\TestFolder" "\.jpg$" true
    Array.iter (fun a -> printfn "%s" a) files

let test03 () =
    let fmt1 = Printf.StringFormat<int->string> "DSCF%04i.jpg" 
    let (opts: UniformRenameParams -> UniformRenameParams) = fun p -> 
        { p with 
            InputFolder = @"G:\work\photos1\TestFolder"
            MatchPattern = "\.jpg$"
            MatchIgnoreCase = true
            MakeName = sprintf fmt1 }
    UniformRename opts



