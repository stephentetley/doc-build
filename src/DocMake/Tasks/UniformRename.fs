[<AutoOpen>]
module DocMake.Tasks.UniformRename

open System.IO
open System.Text.RegularExpressions


open DocMake.Utils.Common

 
// Rename files in a folder matching a pattern. 
// The new names are uniform base name with an numbering.
// The file extension is kept the same.
    

[<CLIMutable>]
type UniformRenameParams = 
    { 
        InputFolder : string
        MatchPattern : string
        MatchIgnoreCase : bool
        NameTemplate : Printf.StringFormat<int->string>
    }


type NameFormat = Printf.StringFormat<int->string>


let UniformRenameDefaults = 
    { InputFolder = "aux"
      MatchPattern = ""
      MatchIgnoreCase = false
      NameTemplate = @"output-%i"}

// Use... System.IO.File.Move (oldname,newname)
// System.IO.Directory.GetFiles(directory)

let matchFiles (dir:string) (search:string) (ignoreCase:bool) : string [] = 
    let re = if ignoreCase then new Regex(search, RegexOptions.IgnoreCase) else new Regex(search)
    Directory.GetFiles(dir) |> Array.filter (fun s -> re.Match(s).Success)



let renameFiles (files:string []) (fmt:NameFormat) : unit = 
    let rename1 count oldpath = 
        let dirname = Path.GetDirectoryName(oldpath)
        let newpath = Path.Combine(dirname, sprintf fmt count)
        printfn "Rename: %s\n    => %s" oldpath newpath
        Directory.Move(oldpath, newpath)
        count+1
    ignore <| Array.fold rename1 1 files
    
    
let UniformRename (setUniformRenameParams: UniformRenameParams -> UniformRenameParams) : unit =
    let opts  = UniformRenameDefaults |> setUniformRenameParams
    let worklist = matchFiles opts.InputFolder opts.MatchPattern opts.MatchIgnoreCase
    renameFiles worklist opts.NameTemplate

    


