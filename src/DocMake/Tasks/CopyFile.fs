[<AutoOpen>]
module DocMake.Tasks.CopyFile

open System.IO


[<CLIMutable>]
type CopyFileParams = 
    { 
        SourceFile : string
        DestFile : string
        Overwrite : bool
        CreateOuputDirectoryIfMissing : bool
    }

let CopyFileDefaults = 
    { SourceFile = @""
      DestFile = @""
      Overwrite = true
      CreateOuputDirectoryIfMissing = true }




let isEmpty (s:string) : bool = 
    match s with 
    | null -> true
    | "" -> true
    | _ -> false


let private copyFile1 (srcPath:string) (destPath:string) (createDir:bool) (overwrite:bool) : unit = 
    try 
        let destFolder = Path.GetDirectoryName destPath
        if not (Directory.Exists destFolder) && createDir then
            ignore <| Directory.CreateDirectory destFolder
        else
            () 
        File.Copy(srcPath, destPath, overwrite)
    with
    | ex -> printfn "CopyFile - Some error occured for %s - '%s'" srcPath ex.Message



let CopyFile (setCopyFileParams: CopyFileParams -> CopyFileParams) : unit =
    let options = CopyFileDefaults |> setCopyFileParams
    if File.Exists(options.SourceFile) && not (isEmpty options.DestFile)
    then
        copyFile1 options.SourceFile options.DestFile options.CreateOuputDirectoryIfMissing options.Overwrite
    else ()