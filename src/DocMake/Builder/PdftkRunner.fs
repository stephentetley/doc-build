// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


module DocMake.Builder.PdftkRunner


open DocMake.Builder.Base

type PdftkEnv = 
    { PdftkExe: string }

/// PandocRunner is an error+reader monad.
type PdftkRunner<'a> = 
    private PdftkRunner of (PdftkEnv -> Result<'a>)
    
let inline private apply1 (ma : PdftkRunner<'a>) (env:PdftkEnv) : Result<'a> = 
    let (PdftkRunner f) = ma in f env


let inline preturn (x:'a) : PdftkRunner<'a> = 
    PdftkRunner (fun _ -> Ok x)

let private failM : PdftkRunner<'a> = 
    PdftkRunner (fun _ -> Err (buildError "failM"))



let inline private bindM (ma:PdftkRunner<'a>) 
                            (f : 'a -> PdftkRunner<'b>) : PdftkRunner<'b> =
    PdftkRunner <| fun env -> 
        match apply1 ma env with
        | Err msg -> Err msg
        | Ok a -> apply1 (f a) env


let inline private altM (ma:PdftkRunner<'a>) 
                            (mb:PdftkRunner<'a>) : PdftkRunner<'a> =
    PdftkRunner <| fun env -> 
        match apply1 ma env with 
        | Err stk1 -> 
            match apply1 mb env with
            | Err stk2 -> Err (concatBuildErrors "altM" [stk1;stk2])
            | Ok b -> Ok b
        | Ok a -> Ok a

/// This is Haskell's (>>).
let inline private combineM (ma:PdftkRunner<unit>) 
                                (mb:PdftkRunner<'b>) : PdftkRunner<'b> = 
    PdftkRunner <| fun env -> 
        match apply1 ma env with
        | Err msg -> Err msg
        | Ok _ -> 
            match apply1 mb env with
            | Err msg -> Err msg
            | Ok b -> Ok b

let inline private delayM (fn:unit -> PdftkRunner<'a>) : PdftkRunner<'a> = 
    bindM (preturn ()) fn 

type PdftkRunnerBuilder() = 
    member self.Return x        = preturn x
    member self.Bind (p,f)      = bindM p f
    member self.Zero ()         = failM
    member self.Delay fn        = delayM fn
    member self.Combine (p,q)   = combineM p q

let (pdftkRunner:PdftkRunnerBuilder) = new PdftkRunnerBuilder()

/// TODO - this should return in the Build monad...
let pdftkRun (env:PdftkEnv) (ma:PdftkRunner<'a>) : Result<'a> = 
    apply1 ma env


