// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Base.Shell

[<AutoOpen>]
module Shell = 

    open System.IO

    type ProcessResult = 
        | ProcSuccess of string
        | ProcErrorCode of int
        | ProcErrorMessage of string

    let exitcodeToResult (code:int) (stdout:string) : ProcessResult = 
        match code with
        | 0 -> ProcSuccess stdout
        | _ -> ProcErrorCode code

    // ************************************************************************
    // RunProcess getting text written to stdout 



    let executeProcess (workingDirectory:string) 
                        (toolPath:string) (command:string) : ProcessResult = 
        try
            use proc = new System.Diagnostics.Process()
            proc.StartInfo.FileName <- toolPath
            proc.StartInfo.Arguments <- command
            proc.StartInfo.WorkingDirectory <- workingDirectory
            proc.StartInfo.UseShellExecute <- false
            proc.StartInfo.RedirectStandardOutput <- true
            proc.Start() |> ignore

            let reader : System.IO.StreamReader = proc.StandardOutput
            let stdout = reader.ReadToEnd()

            proc.WaitForExit () 
            exitcodeToResult proc.ExitCode stdout
        with
        | ex -> ProcErrorMessage (sprintf "executeProcess: \n%s" ex.Message)