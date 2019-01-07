// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


namespace DocBuild.Base

[<AutoOpen>]
module Common = 

    open System.IO






    

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

    // ************************************************************************
    // RunProcess = 

    /// Return `Choice2Of2 0` indicates Success
    let executeProcess (workingDirectory:string) (toolPath:string) (command:string) : Choice<string,int> = 
        try
            let procInfo = new System.Diagnostics.ProcessStartInfo ()
            procInfo.WorkingDirectory <- workingDirectory
            procInfo.FileName <- toolPath
            procInfo.Arguments <- command
            procInfo.CreateNoWindow <- true
            let proc = new System.Diagnostics.Process()
            proc.StartInfo <- procInfo
            proc.Start() |> ignore
            proc.WaitForExit () 
            Choice2Of2 <| proc.ExitCode
        with
        | ex -> Choice1Of2 (sprintf "executeProcess: \n%s" ex.Message)