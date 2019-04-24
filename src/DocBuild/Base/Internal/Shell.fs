// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Base.Internal

/// TODO - this can be removed when the SL-Format dependency is updated.

module Shell = 

    open System.IO

    let private procErrorCodeMessage (code:int) : string = 
        sprintf "process error: code %i" code

    let private procErrorMessage (msg:string) : string = 
        sprintf "process error: '%s'" msg
        


    let exitcodeToResult (errCode:int) (stdOutput:string) : Result<string,string> = 
        match errCode with
        | 0 -> Ok stdOutput
        | _ -> Error (procErrorCodeMessage errCode)

    type ProcessOptions = 
        { WorkingDirectory: string 
          ExecutableName: string 
        }

    // ************************************************************************
    // RunProcess getting text written to stdout 

    /// Result is contents of stdout
    let executeProcess (procOptions:ProcessOptions) 
                        (command:string) : Result<string, string> = 
        try
            use proc = new System.Diagnostics.Process()
            proc.StartInfo.FileName <- procOptions.ExecutableName
            proc.StartInfo.Arguments <- command
            proc.StartInfo.WorkingDirectory <- procOptions.WorkingDirectory
            proc.StartInfo.UseShellExecute <- false
            proc.StartInfo.RedirectStandardOutput <- true
            proc.Start() |> ignore

            let reader : System.IO.StreamReader = proc.StandardOutput
            let stdout = reader.ReadToEnd()

            proc.WaitForExit () 
            exitcodeToResult proc.ExitCode stdout
        with
        | ex -> Error (procErrorMessage (sprintf "executeProcess: \n%s" ex.Message))



