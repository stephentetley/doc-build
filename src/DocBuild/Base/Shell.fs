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

    type ProcessOptions = 
        { WorkingDirectory: string 
          ExecutableName: string 
        }

    // ************************************************************************
    // RunProcess getting text written to stdout 

    let executeProcess (procOptions:ProcessOptions) (command:string) : ProcessResult = 
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
        | ex -> ProcErrorMessage (sprintf "executeProcess: \n%s" ex.Message)

    // ************************************************************************
    // Options


    /// Note - name is expected to contain "-" or "--"
    type CommandArg = 
        | NoArg of string
        | ReqArg of string * string
        member x.Option 
            with get() : string = 
                match x with
                | NoArg name -> name
                | ReqArg(name,value) -> sprintf "%s=%s" name value

    type CommandArgs = 
        val private Commands : CommandArg list

        new (args:CommandArg list) = 
            { Commands = args }
        
        new (arg:CommandArg) = 
            { Commands = [arg] }

        static member (^^) (x:CommandArgs, y:CommandArgs) : CommandArgs = 
            new CommandArgs(args = x.Commands @ y.Commands)

        member x.Command 
            with get() : string = 
                x.Commands |> List.map (fun x -> x.Option) |> String.concat " "

        static member Concat(xs:CommandArgs list) : CommandArgs = 
            let xss = xs |> List.map (fun x -> x.Commands)
            new CommandArgs(args = List.concat xss)

    let reqArg (name:string) (value:string) : CommandArgs = 
        new CommandArgs(arg = ReqArg(name,value))

    let noArg (name:string) : CommandArgs = 
        new CommandArgs(arg = NoArg(name))

    let emptyArgs : CommandArgs = 
        new CommandArgs(args = [])