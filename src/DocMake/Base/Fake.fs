module DocMake.Base.Fake

open Fake
open Fake.Core
open Fake.Core.Globbing.Operators

    
let private tryOneOrMore (input:'a list) : option<'a list> = 
    match input with
    | [] -> None
    | _ -> Some input

let private tryExactlyOne (input:'a list) : option<'a> = 
    match input with
    | [x] -> Some x
    | _ -> None

//let private tryAtMostOne (input:'a list) : option<'a> = 
//    match input with
//    | [] -> ???
//    | [x] -> Some x
//    | _ -> None

// Zero or more matches.
// No need for a try variant (empty list is no matches)
// Note - pattern is a glob, not a regex.
let findAllMatchingFiles (pattern:string) (dir:string) : string list = 
    !! (dir @@ pattern) |> Seq.toList
    
// One or more matches. 
// Note - pattern is a glob, not a regex.
let tryFindSomeMatchingFiles (pattern:string) (dir:string) : option<string list> = 
    !! (dir @@ pattern) |> Seq.toList |> tryOneOrMore

// Exactly one matches.
// Note - pattern is a glob, not a regex.
let tryFindExactlyOneMatchingFile (pattern:string) (dir:string) : option<string> = 
    !! (dir @@ pattern) |> Seq.toList |> tryExactlyOne

// We have tried the following combinators but they seem to be less
// clear in user code than using match ... with
//let optionMandatory (source:'a option) (failMsg:string) (success:'a -> unit) = 
//    match source with
//    | Some a -> success a
//    | None -> failwith failMsg

//let optionOptional (source:'a option) (warnMsg:string) (success:'a -> unit) = 
//    match source with
//    | Some a -> success a
//    | None -> Trace.tracefn "%s" warnMsg

let assertMandatory (failMsg:string) : unit = failwithf "FAIL: Mandatory: %s" failMsg

let assertOptional  (warnMsg:string) : unit = Trace.tracefn "WARN: Optional: %s" warnMsg

