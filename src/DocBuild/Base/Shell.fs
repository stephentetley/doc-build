// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Base.Shell

[<AutoOpen>]
module Shell = 

    open System.IO

    type ProcessResult = 
        | ProcSuccess
        | ProcErrorCode of int
        | ProcErrorMessage of string

    let exitcodeToResult (code:int) : ProcessResult = 
        match code with
        | 0 -> ProcSuccess
        | _ -> ProcErrorCode code

    // ************************************************************************
    // RunProcess = 

    /// Return `Choice2Of2 0` indicates Success
    let executeProcess (workingDirectory:string) 
                        (toolPath:string) (command:string) : ProcessResult = 
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
            exitcodeToResult <| proc.ExitCode
        with
        | ex -> ProcErrorMessage (sprintf "executeProcess: \n%s" ex.Message)