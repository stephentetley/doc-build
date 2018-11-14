// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


module DocMake.Builder.PowerPointRunner

open Microsoft.Office.Interop

open DocMake.Base.OfficeUtils
open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis

type PowerPointHandle = 
    
    val mutable PowerPointApp:PowerPoint.Application

    new () = 
        { PowerPointApp = initPowerPoint () }

    member v.CleanUp () = 
        finalizePowerPoint v.PowerPointApp

    member v.Handle 
        /// Word has a problem with dying during long tasks
        /// that open and close lots of files
        with get() : PowerPoint.Application = 
                match v.PowerPointApp with
                | null -> 
                    let pptApp1 = initPowerPoint ()
                    v.PowerPointApp <- pptApp1
                    pptApp1
                | app -> app


    


/// PowerPointRunner is reader over BuildMonad.
type PowerPointRunner<'a> = 
    private PowerPointRunner of (PowerPointHandle -> BuildMonad<unit,'a>)
    
let inline private apply1 (ma : PowerPointRunner<'a>) 
                            (res:PowerPointHandle) : BuildMonad<unit,'a> = 
    let (PowerPointRunner f) = ma in f res


let inline mreturn (x:'a) : PowerPointRunner<'a> = 
    PowerPointRunner (fun _ -> breturn x)

let private failM : PowerPointRunner<'a> = 
    PowerPointRunner (fun _ -> throwError "PowerPointRunner - failM")



let inline private bindM (ma:PowerPointRunner<'a>) 
                            (f : 'a -> PowerPointRunner<'b>) : PowerPointRunner<'b> =
    PowerPointRunner <| fun env -> 
        buildMonad { 
            let! a = apply1 ma env
            let! b = apply1 (f a) env
            return b
        }


let inline private altM (ma:PowerPointRunner<'a>) 
                            (mb:PowerPointRunner<'a>) : PowerPointRunner<'a> =
    PowerPointRunner <| fun env -> 
        apply1 ma env <|> apply1 mb env

/// This is Haskell's (>>).
let inline private combineM (ma:PowerPointRunner<unit>) 
                                (mb:PowerPointRunner<'b>) : PowerPointRunner<'b> = 
    PowerPointRunner <| fun env -> 
        apply1 ma env >>. apply1 mb env

let inline private delayM (fn:unit -> PowerPointRunner<'a>) : PowerPointRunner<'a> = 
    bindM (mreturn ()) fn 

type PowerPointRunnerBuilder() = 
    member self.Return x        = mreturn x
    member self.Bind (p,f)      = bindM p f
    member self.Zero ()         = failM
    member self.Delay fn        = delayM fn
    member self.Combine (p,q)   = combineM p q

let (powerPointRunner:PowerPointRunnerBuilder) = new PowerPointRunnerBuilder()


let powerPointRun (ma:PowerPointRunner<'a>) : BuildMonad<unit,'a> = 
    let handle = new PowerPointHandle ()
    let ans = apply1 ma handle
    handle.CleanUp ()
    ans


let liftBM (ma:BuildMonad<unit,'a>) : PowerPointRunner<'a> =  
    PowerPointRunner <| fun _ -> ma

let powerPointExec (operation:PowerPoint.Application -> 'a) : PowerPointRunner<'a> = 
    PowerPointRunner <| fun res -> 
        let handle = res.Handle
        let ans = operation handle 
        breturn ans
