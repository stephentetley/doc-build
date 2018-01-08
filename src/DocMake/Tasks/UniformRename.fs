[<AutoOpen>]
module DocMake.Tasks.UniformRename

open System.IO
open System.Text.RegularExpressions


open DocMake.Base.Common

 
// Rename files in a folder matching a pattern. 
// The new names are uniform base name with an numbering.
// The file extension is kept the same.

// Push whether or not to use sprintf to the client, this makes things more flexible
type NameFormatter = int->string

[<CLIMutable>]
type UniformRenameParams = 
    { 
        InputFolder : string
        MatchPattern : string
        MatchIgnoreCase : bool
        MakeName : NameFormatter
    }





let UniformRenameDefaults = 
    { InputFolder = "aux"
      MatchPattern = ""
      MatchIgnoreCase = false
      MakeName = sprintf @"output-%i"}

// Use... System.IO.File.Move (oldname,newname)
// System.IO.Directory.GetFiles(directory)

let matchFiles (dir:string) (search:string) (ignoreCase:bool) : string [] = 
    let re = if ignoreCase then new Regex(search, RegexOptions.IgnoreCase) else new Regex(search)
    Directory.GetFiles(dir) |> Array.filter (fun s -> re.Match(s).Success)



let renameFiles (files:string []) (fmt:NameFormatter) : unit = 
    let rename1 count oldpath = 
        let dirname = Path.GetDirectoryName(oldpath)
        let newpath = Path.Combine(dirname, fmt count)
        printfn "Rename: %s\n    => %s" oldpath newpath
        Directory.Move(oldpath, newpath)
        count+1
    ignore <| Array.fold rename1 1 files
    
    
let UniformRename (setUniformRenameParams: UniformRenameParams -> UniformRenameParams) : unit =
    let opts  = UniformRenameDefaults |> setUniformRenameParams
    let worklist = matchFiles opts.InputFolder opts.MatchPattern opts.MatchIgnoreCase
    renameFiles worklist opts.MakeName

    


