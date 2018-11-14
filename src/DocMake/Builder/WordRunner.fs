// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


module DocMake.Builder.WordRunner

open Microsoft.Office.Interop

open DocMake.Base.OfficeUtils
open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis

type WordHandle = 
    
    val mutable WordApp:Word.Application

    new () = 
        { WordApp = initWord () }

    member v.CleanUp () = 
        finalizeWord v.WordApp

    member v.Handle 
        /// Word has a problem with dying during long tasks
        /// that open and close lots of files
        with get() : Word.Application = 
                match v.WordApp with
                | null -> 
                    let wordApp1 = initWord ()
                    v.WordApp <- wordApp1
                    wordApp1
                | app -> app


    


/// WordRunner is reader over BuildMonad.
type WordRunner<'a> = 
    private WordRunner of (WordHandle -> BuildMonad<unit,'a>)
    
let inline private apply1 (ma : WordRunner<'a>) 
                            (res:WordHandle) : BuildMonad<unit,'a> = 
    let (WordRunner f) = ma in f res


let inline mreturn (x:'a) : WordRunner<'a> = 
    WordRunner (fun _ -> breturn x)

let private failM : WordRunner<'a> = 
    WordRunner (fun _ -> throwError "WordRunner - failM")



let inline private bindM (ma:WordRunner<'a>) 
                            (f : 'a -> WordRunner<'b>) : WordRunner<'b> =
    WordRunner <| fun env -> 
        buildMonad { 
            let! a = apply1 ma env
            let! b = apply1 (f a) env
            return b
        }


let inline private altM (ma:WordRunner<'a>) 
                            (mb:WordRunner<'a>) : WordRunner<'a> =
    WordRunner <| fun env -> 
        apply1 ma env <|> apply1 mb env

/// This is Haskell's (>>).
let inline private combineM (ma:WordRunner<unit>) 
                                (mb:WordRunner<'b>) : WordRunner<'b> = 
    WordRunner <| fun env -> 
        apply1 ma env >>. apply1 mb env

let inline private delayM (fn:unit -> WordRunner<'a>) : WordRunner<'a> = 
    bindM (mreturn ()) fn 

type WordRunnerBuilder() = 
    member self.Return x        = mreturn x
    member self.Bind (p,f)      = bindM p f
    member self.Zero ()         = failM
    member self.Delay fn        = delayM fn
    member self.Combine (p,q)   = combineM p q

let (wordRunner:WordRunnerBuilder) = new WordRunnerBuilder()


let wordRun (ma:WordRunner<'a>) : BuildMonad<unit,'a> = 
    let handle = new WordHandle ()
    let ans = apply1 ma handle
    handle.CleanUp ()
    ans


let liftBM (ma:BuildMonad<unit,'a>) : WordRunner<'a> =  
    WordRunner <| fun _ -> ma

let wordExec (operation:Word.Application -> 'a) : WordRunner<'a> = 
    WordRunner <| fun res -> 
        let handle = res.Handle
        let ans = operation handle 
        breturn ans
