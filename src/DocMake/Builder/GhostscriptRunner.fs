// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


module DocMake.Builder.GhostscriptRunner


open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis

type GhostscriptEnv = 
    { GhostscriptExe: string }

/// GhostscriptRunner is reader over BuildMonad.
type GhostscriptRunner<'a> = 
    private GhostscriptRunner of (GhostscriptEnv -> BuildMonad<unit,'a>)
    
let inline private apply1 (ma : GhostscriptRunner<'a>) 
                            (env:GhostscriptEnv) : BuildMonad<unit,'a> = 
    let (GhostscriptRunner f) = ma in f env


let inline mreturn (x:'a) : GhostscriptRunner<'a> = 
    GhostscriptRunner (fun _ -> breturn x)

let private failM : GhostscriptRunner<'a> = 
    GhostscriptRunner (fun _ -> throwError "GhostscriptRunner - failM")



let inline private bindM (ma:GhostscriptRunner<'a>) 
                            (f : 'a -> GhostscriptRunner<'b>) : GhostscriptRunner<'b> =
    GhostscriptRunner <| fun env -> 
        buildMonad { 
            let! a = apply1 ma env
            let! b = apply1 (f a) env
            return b
        }


let inline private altM (ma:GhostscriptRunner<'a>) 
                            (mb:GhostscriptRunner<'a>) : GhostscriptRunner<'a> =
    GhostscriptRunner <| fun env -> 
        apply1 ma env <|> apply1 mb env

/// This is Haskell's (>>).
let inline private combineM (ma:GhostscriptRunner<unit>) 
                                (mb:GhostscriptRunner<'b>) : GhostscriptRunner<'b> = 
    GhostscriptRunner <| fun env -> 
        apply1 ma env >>. apply1 mb env

let inline private delayM (fn:unit -> GhostscriptRunner<'a>) : GhostscriptRunner<'a> = 
    bindM (mreturn ()) fn 

type GhostscriptRunnerBuilder() = 
    member self.Return x        = mreturn x
    member self.Bind (p,f)      = bindM p f
    member self.Zero ()         = failM
    member self.Delay fn        = delayM fn
    member self.Combine (p,q)   = combineM p q

let (ghostscriptRunner:GhostscriptRunnerBuilder) = new GhostscriptRunnerBuilder()


let ghostscriptRun (env:GhostscriptEnv) (ma:GhostscriptRunner<'a>) : BuildMonad<unit,'a> = 
    apply1 ma env

let liftBM (ma:BuildMonad<unit,'a>) : GhostscriptRunner<'a> =  
    GhostscriptRunner <| fun _ -> ma

let ghostscriptExec (command:string) : GhostscriptRunner<unit> = 
    GhostscriptRunner <| fun env -> 
        shellRun env.GhostscriptExe command "Pdftk failed"
