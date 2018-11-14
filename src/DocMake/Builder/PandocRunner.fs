// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


module DocMake.Builder.PandocRunner

open MarkdownDoc
open MarkdownDoc.Pandoc

open DocMake.Builder.BuildMonad
open DocMake.Builder.Document

type PandocEnv = 
    { WorkingDirectory: string
      DocxReferenceDoc: string option }

/// PandocRunner is reader over BuildMonad.
type PandocRunner<'a> = 
    private PandocRunner of (PandocEnv -> BuildMonad<unit,'a>)
    
let inline private apply1 (ma : PandocRunner<'a>) 
                            (env:PandocEnv) : BuildMonad<unit,'a> = 
    let (PandocRunner f) = ma in f env


let inline mreturn (x:'a) : PandocRunner<'a> = 
    PandocRunner (fun _ -> breturn x)

let private failM : PandocRunner<'a> = 
    PandocRunner (fun _ -> throwError "PandocRunner - failM")



let inline private bindM (ma:PandocRunner<'a>) (f : 'a -> PandocRunner<'b>) : PandocRunner<'b> =
    PandocRunner <| fun env -> 
        buildMonad { 
            let! a = apply1 ma env
            let! b = apply1 (f a) env
            return b
        }



let inline private altM (ma:PandocRunner<'a>) (mb:PandocRunner<'a>) : PandocRunner<'a> =
    PandocRunner <| fun env -> 
        apply1 ma env <|> apply1 mb env

/// This is Haskell's (>>).
let inline private combineM (ma:PandocRunner<unit>) (mb:PandocRunner<'b>) : PandocRunner<'b> = 
    PandocRunner <| fun env -> 
        apply1 ma env >>. apply1 mb env
        

let inline private delayM (fn:unit -> PandocRunner<'a>) : PandocRunner<'a> = 
    bindM (mreturn ()) fn 

type PandocRunnerBuilder() = 
    member self.Return x        = mreturn x
    member self.Bind (p,f)      = bindM p f
    member self.Zero ()         = failM
    member self.Delay fn        = delayM fn
    member self.Combine (p,q)   = combineM p q

let (pandocRunner:PandocRunnerBuilder) = new PandocRunnerBuilder()


let pandocRun (env:PandocEnv) (ma:PandocRunner<'a>) : BuildMonad<unit,'a> = 
    apply1 ma env

let liftBM (ma:BuildMonad<unit,'a>) : PandocRunner<'a> =  
    PandocRunner <| fun _ -> ma

/// TODO - MarkdownDoc.PandocInvoke should probably return error code 
/// on failure rather than throwing an exception.
let generateDocx (doc:Markdown) (outputPath:string) (otherOptions: PandocOption list) : PandocRunner<WordDoc> = 
    PandocRunner <| fun env ->
        let stylesDoc = Option.defaultValue "" env.DocxReferenceDoc
        pandocGenerateDocx env.WorkingDirectory doc stylesDoc outputPath otherOptions
        breturn (Document.makeDocument outputPath)


