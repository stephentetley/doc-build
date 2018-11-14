// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

module DocMake.Builder.Base

open System.Text




type BuildError = BuildError of string * list<BuildError>

let getErrorLog (err:BuildError) : string = 
    let writeLine (depth:int) (str:string) (sb:StringBuilder) : StringBuilder = 
        let line = sprintf "%s %s" (String.replicate depth "*") str
        sb.AppendLine(line)
    let rec work (e1:BuildError) (depth:int) (sb:StringBuilder) : StringBuilder  = 
        match e1 with
        | BuildError (s,[]) -> writeLine depth s sb
        | BuildError (s,xs) ->
            let sb1 = writeLine depth s sb
            List.fold (fun buf branch -> work branch (depth+1) buf) sb1 xs
    work err 0 (new StringBuilder()) |> fun sb -> sb.ToString()

/// Create a BuildError
let buildError (errMsg:string) : BuildError = 
    BuildError(errMsg, [])

let concatBuildErrors (errMsg:string) (failures:BuildError list) : BuildError = 
    BuildError(errMsg,failures)
    


type Result<'a> =
    | Err of BuildError
    | Ok of 'a