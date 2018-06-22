// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

module DocMake.Base.FakeExtras

open Fake.Core
open Fake.Core.Globbing.Operators

open DocMake.Base.FakeFake
    
let private tryOneOrMore (input:'a list) : option<'a list> = 
    match input with
    | [] -> None
    | _ -> Some input

let private tryExactlyOne (input:'a list) : option<'a> = 
    match input with
    | [x] -> Some x
    | _ -> None

        
// Has one or more matches. 
// Note - pattern is a glob, not a regex.
let hasMatchingFiles (pattern:string) (dir:string) : bool = 
    let test = not << Seq.isEmpty
    !! (dir </> pattern) |> test


// Zero or more matches.
// No need for a try variant (empty list is no matches)
// Note - pattern is a glob, not a regex.
let findAllMatchingFiles (pattern:string) (dir:string) : string list = 
    !! (dir </> pattern) |> Seq.toList
    
// One or more matches. 
// Note - pattern is a glob, not a regex.
let tryFindSomeMatchingFiles (pattern:string) (dir:string) : option<string list> = 
    !! (dir </> pattern) |> Seq.toList |> tryOneOrMore

// Exactly one matches.
// Note - pattern is a glob, not a regex.
let tryFindExactlyOneMatchingFile (pattern:string) (dir:string) : option<string> = 
    !! (dir </> pattern) |> Seq.toList |> tryExactlyOne


let assertMandatory (failMsg:string) : unit = failwithf "FAIL: Mandatory: %s" failMsg

let assertOptional  (warnMsg:string) : unit = Trace.tracefn "WARN: Optional: %s" warnMsg

let subdirectoriesWithMatches (pattern:string) (dir:string) : string list = 
    let dirs = System.IO.Directory.GetDirectories(dir) |> Array.toList
    List.filter (hasMatchingFiles pattern) dirs