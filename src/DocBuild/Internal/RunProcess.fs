// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

namespace DocBuild.Internal.RunProcess

[<AutoOpen>]
module RunProcess = 

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