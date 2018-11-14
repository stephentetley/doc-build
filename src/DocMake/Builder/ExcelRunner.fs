// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


module DocMake.Builder.ExcelRunner

open Microsoft.Office.Interop

open DocMake.Base.OfficeUtils
open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis

type ExcelHandle = 
    
    val mutable ExcelApp:Excel.Application

    new () = 
        { ExcelApp = initExcel () }

    member v.CleanUp () = 
        finalizeExcel v.ExcelApp

    member v.Handle 
        with get() : Excel.Application = 
                match v.ExcelApp with
                | null -> 
                    let excel1 = initExcel ()
                    v.ExcelApp <- excel1
                    excel1
                | app -> app


    


/// ExcelRunner is reader over BuildMonad.
type ExcelRunner<'a> = 
    private ExcelRunner of (ExcelHandle -> BuildMonad<unit,'a>)
    
let inline private apply1 (ma : ExcelRunner<'a>) 
                            (res:ExcelHandle) : BuildMonad<unit,'a> = 
    let (ExcelRunner f) = ma in f res


let inline mreturn (x:'a) : ExcelRunner<'a> = 
    ExcelRunner (fun _ -> breturn x)

let private failM : ExcelRunner<'a> = 
    ExcelRunner (fun _ -> throwError "ExcelRunner - failM")



let inline private bindM (ma:ExcelRunner<'a>) 
                            (f : 'a -> ExcelRunner<'b>) : ExcelRunner<'b> =
    ExcelRunner <| fun env -> 
        buildMonad { 
            let! a = apply1 ma env
            let! b = apply1 (f a) env
            return b
        }


let inline private altM (ma:ExcelRunner<'a>) 
                            (mb:ExcelRunner<'a>) : ExcelRunner<'a> =
    ExcelRunner <| fun env -> 
        apply1 ma env <|> apply1 mb env

/// This is Haskell's (>>).
let inline private combineM (ma:ExcelRunner<unit>) 
                                (mb:ExcelRunner<'b>) : ExcelRunner<'b> = 
    ExcelRunner <| fun env -> 
        apply1 ma env >>. apply1 mb env

let inline private delayM (fn:unit -> ExcelRunner<'a>) : ExcelRunner<'a> = 
    bindM (mreturn ()) fn 

type ExcelRunnerBuilder() = 
    member self.Return x        = mreturn x
    member self.Bind (p,f)      = bindM p f
    member self.Zero ()         = failM
    member self.Delay fn        = delayM fn
    member self.Combine (p,q)   = combineM p q

let (excelRunner:ExcelRunnerBuilder) = new ExcelRunnerBuilder()


let excelRun (ma:ExcelRunner<'a>) : BuildMonad<unit,'a> = 
    let handle = new ExcelHandle ()
    let ans = apply1 ma handle
    handle.CleanUp ()
    ans


let liftBM (ma:BuildMonad<unit,'a>) : ExcelRunner<'a> =  
    ExcelRunner <| fun _ -> ma

let excelExec (operation:Excel.Application -> 'a) : ExcelRunner<'a> = 
    ExcelRunner <| fun res -> 
        let handle = res.Handle
        let ans = operation handle 
        breturn ans
