// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


module DocMake.Builder.PandocRunner

open MarkdownDoc
open MarkdownDoc.Pandoc

type Result<'a> =
    | Err of string
    | Ok of 'a

type PandocEnv = 
    { WorkingDirectory: string
      DocxReferenceDoc: string option }

/// PandocRunner is an error+reader monad.
type PandocRunner<'a> = 
    private PandocRunner of (PandocEnv -> Result<'a>)
    
let inline private apply1 (ma : PandocRunner<'a>) (env:PandocEnv) : Result<'a> = 
    let (PandocRunner f) = ma in f env

// Return in the BuildMonad
let inline preturn (x:'a) : PandocRunner<'a> = 
    PandocRunner (fun _ -> Ok x)

let private failM : PandocRunner<'a> = 
    PandocRunner (fun _ -> Err "failM")



let inline private bindM (ma:PandocRunner<'a>) (f : 'a -> PandocRunner<'b>) : PandocRunner<'b> =
    PandocRunner <| fun env -> 
        match apply1 ma env with
        | Err msg -> Err msg
        | Ok a -> apply1 (f a) env


let inline private altM (ma:PandocRunner<'a>) (mb:PandocRunner<'a>) : PandocRunner<'a> =
    PandocRunner <| fun env -> 
        match apply1 ma env with 
        | Err _ -> 
            match apply1 mb env with
            | Err _ -> Err "altM"
            | Ok b -> Ok b
        | Ok a -> Ok a

/// This is Haskell's (>>).
let inline private combineM (ma:PandocRunner<unit>) (mb:PandocRunner<'b>) : PandocRunner<'b> = 
    PandocRunner <| fun env -> 
        match apply1 ma env with
        | Err msg -> Err msg
        | Ok _ -> 
            match apply1 mb env with
            | Err msg -> Err msg
            | Ok b -> Ok b

let inline private delayM (fn:unit -> PandocRunner<'a>) : PandocRunner<'a> = 
    bindM (preturn ()) fn 

type PandocRunnerBuilder() = 
    member self.Return x        = preturn x
    member self.Bind (p,f)      = bindM p f
    member self.Zero ()         = failM
    member self.Delay fn        = delayM fn
    member self.Combine (p,q)   = combineM p q

let (pandocRunner:PandocRunnerBuilder) = new PandocRunnerBuilder()

/// TODO - this should return in the Build monad...
let pandocRun (env:PandocEnv) (ma:PandocRunner<'a>) : Result<'a> = 
    apply1 ma env

/// TODO - this should return a Document<Docx>
/// TODO - MarkdownDoc.PandocInvoke should probably return error 
/// code on failure rather than throwing an exception.
let generateDocx (doc:Markdown) (outputPath:string) (otherOptions: PandocOption list) : PandocRunner<unit> = 
    PandocRunner <| fun env ->
        let stylesDoc = Option.defaultValue "" env.DocxReferenceDoc
        pandocGenerateDocx env.WorkingDirectory doc stylesDoc outputPath otherOptions
        Ok ()


