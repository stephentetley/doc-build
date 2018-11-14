// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


module DocMake.Builder.PdftkRunner


open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis

type PdftkEnv = 
    { PdftkExe: string }

/// PdftkRunner is reader over BuildMonad.
type PdftkRunner<'a> = 
    private PdftkRunner of (PdftkEnv -> BuildMonad<unit,'a>)
    
let inline private apply1 (ma : PdftkRunner<'a>) 
                            (env:PdftkEnv) : BuildMonad<unit,'a> = 
    let (PdftkRunner f) = ma in f env


let inline mreturn (x:'a) : PdftkRunner<'a> = 
    PdftkRunner (fun _ -> breturn x)

let private failM : PdftkRunner<'a> = 
    PdftkRunner (fun _ -> throwError "PdftkRunner - failM")



let inline private bindM (ma:PdftkRunner<'a>) 
                            (f : 'a -> PdftkRunner<'b>) : PdftkRunner<'b> =
    PdftkRunner <| fun env -> 
        buildMonad { 
            let! a = apply1 ma env
            let! b = apply1 (f a) env
            return b
        }


let inline private altM (ma:PdftkRunner<'a>) 
                            (mb:PdftkRunner<'a>) : PdftkRunner<'a> =
    PdftkRunner <| fun env -> 
        apply1 ma env <|> apply1 mb env

/// This is Haskell's (>>).
let inline private combineM (ma:PdftkRunner<unit>) 
                                (mb:PdftkRunner<'b>) : PdftkRunner<'b> = 
    PdftkRunner <| fun env -> 
        apply1 ma env >>. apply1 mb env

let inline private delayM (fn:unit -> PdftkRunner<'a>) : PdftkRunner<'a> = 
    bindM (mreturn ()) fn 

type PdftkRunnerBuilder() = 
    member self.Return x        = mreturn x
    member self.Bind (p,f)      = bindM p f
    member self.Zero ()         = failM
    member self.Delay fn        = delayM fn
    member self.Combine (p,q)   = combineM p q

let (pdftkRunner:PdftkRunnerBuilder) = new PdftkRunnerBuilder()


let pdftkRun (env:PdftkEnv) (ma:PdftkRunner<'a>) : BuildMonad<unit,'a> = 
    apply1 ma env

let liftBM (ma:BuildMonad<unit,'a>) : PdftkRunner<'a> =  
    PdftkRunner <| fun _ -> ma


let pdftkExec (command:string) : PdftkRunner<unit> = 
    PdftkRunner <| fun env -> 
        shellRun env.PdftkExe command "Pdftk failed"
