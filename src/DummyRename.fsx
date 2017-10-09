
// For Fsi, remember to load the "trail" of modules not just the module you want, otherwise
// you get athe error: "The namespace or module 'Xyz' is not defined".

#load @"DocMake\Utils\Common.fs"
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



