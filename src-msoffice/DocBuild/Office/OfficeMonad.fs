// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Office.OfficeMonad


[<AutoOpen>]
module OfficeMonad = 

    open Microsoft.Office.Interop
    open DocBuild.Base

    open DocBuild.Office.Internal


    type OfficeHandles = 
        val mutable private WordHandle : Word.Application 
        val mutable private ExcelHandle : Excel.Application
        val mutable private PowerPointHandle : PowerPoint.Application

        new () = 
            { WordHandle = null
              ExcelHandle = null
              PowerPointHandle = null }

        /// Opens a handle as needed.
        member x.WordApp : Word.Application = 
            match x.WordHandle with
            | null -> 
                let word1 = initWord ()
                x.WordHandle <- word1
                word1
            | app -> app

        /// Opens a handle as needed.
        member x.ExcelApp : Excel.Application  = 
            match x.ExcelHandle with
            | null -> 
                let excel1 = initExcel ()
                x.ExcelHandle <- excel1
                excel1
            | app -> app

        /// Opens a handler as needed.
        member x.PowerPointApp : PowerPoint.Application  = 
            match x.PowerPointHandle with
            | null -> 
                let powerPoint1 = initPowerPoint ()
                x.PowerPointHandle <- powerPoint1
                powerPoint1
            | app -> app

        
        member x.RunFinalizer () = 
            match x.WordHandle with
            | null -> () 
            | word -> finalizeWord word
            match x.ExcelHandle with
            | null -> ()
            | excel -> finalizeExcel excel
            match x.PowerPointHandle with
            | null -> ()
            | ppt -> finalizePowerPoint ppt

    type OfficeMonad<'a> = 
        OfficeMonad of (OfficeHandles -> DocMonad.DocMonad<'a>)

    let inline private apply1 (ma: OfficeMonad<'a>) 
                                (env: OfficeHandles) : DocMonad.DocMonad<'a> = 
        let (OfficeMonad f) = ma in f env

    let inline oreturn (x:'a) : OfficeMonad<'a> = 
        OfficeMonad <| fun _ -> DocMonad.breturn x

    let inline bindM (ma:OfficeMonad<'a>) 
                        (f :'a -> OfficeMonad<'b>) : OfficeMonad<'b> =
        OfficeMonad <| fun env -> 
            DocMonad.docMonad.Bind(apply1 ma env, fun a -> apply1 (f a) env)

    let inline ozero () : OfficeMonad<'a> = 
        OfficeMonad <| fun _ -> DocMonad.bzero () 

    /// "First success"
    let inline private combineM (ma:OfficeMonad<'a>) 
                                    (mb:OfficeMonad<'a>) : OfficeMonad<'a> = 
        OfficeMonad <| fun env -> 
            DocMonad.docMonad.Combine(apply1 ma env, apply1 mb env)


    let inline private  delayM (fn:unit -> OfficeMonad<'a>) : OfficeMonad<'a> = 
        bindM (oreturn ()) fn 

    type OfficeMonadBuilder() = 
        member self.Return x            = oreturn x
        member self.Bind (p,f)          = bindM p f
        member self.Zero ()             = ozero ()
        member self.Combine (ma,mb)     = combineM ma mb
        member self.Delay fn            = delayM fn

    let (officeMonad:OfficeMonadBuilder) = new OfficeMonadBuilder()


    // ****************************************************
    // Run
    let runOfficeMonad (config:DocMonad.BuilderEnv) 
                    (ma:OfficeMonad<'a>) : BuildResult<'a> = 
        let handles = new OfficeHandles()
        let (OfficeMonad mf) = ma

        let ans = DocMonad.runDocMonad config (mf handles)
        handles.RunFinalizer ()
        ans


    let execOfficeMonad (config:DocMonad.BuilderEnv) 
                        (ma:OfficeMonad<'a>) : 'a = 
        match runOfficeMonad config ma with
        | Ok a -> a
        | Error msg -> failwith msg

    // ****************************************************
    // Lift

    let liftDocMonad (ma:DocMonad.DocMonad<'a>) : OfficeMonad<'a> = 
        OfficeMonad <| fun _ -> ma



    // ****************************************************
    // Errors

    let throwError (msg:string) : OfficeMonad<'a> = 
        liftDocMonad <| DocMonad.throwError msg

    let swapError (msg:string) (ma:OfficeMonad<'a>) : OfficeMonad<'a> = 
        OfficeMonad <| fun env -> 
            DocMonad.swapError msg (apply1 ma env)


    ///// Execute an action that may throw an exception.
    ///// Capture the exception with try ... with
    ///// and return the answer or the expection message in the monad.
    let attempt (ma: OfficeMonad<'a>) : OfficeMonad<'a> = 
        OfficeMonad <| fun env -> 
            DocMonad.attempt (apply1 ma env)



    // ****************************************************
    // Monadic operations


    /// fmap 
    let fmapM (fn:'a -> 'b) (ma:OfficeMonad<'a>) : OfficeMonad<'b> = 
        OfficeMonad <| fun env -> 
           DocMonad.fmapM fn  (apply1 ma env)


    // liftM (which is fmap)
    let liftM (fn:'a -> 'x) (ma:OfficeMonad<'a>) : OfficeMonad<'x> = 
        fmapM fn ma


    let liftM2 (fn:'a -> 'b -> 'x) 
                (ma:OfficeMonad<'a>) 
                (mb:OfficeMonad<'b>) : OfficeMonad<'x> = 
        officeMonad { 
            let! a = ma
            let! b = mb
            return (fn a b)
        }

    let liftM3 (fn:'a -> 'b -> 'c -> 'x) 
                (ma:OfficeMonad<'a>) 
                (mb:OfficeMonad<'b>) 
                (mc:OfficeMonad<'c>) : OfficeMonad<'x> = 
        officeMonad { 
            let! a = ma
            let! b = mb
            let! c = mc
            return (fn a b c)
        }

    let liftM4 (fn:'a -> 'b -> 'c -> 'd -> 'x) 
                (ma:OfficeMonad<'a>) 
                (mb:OfficeMonad<'b>) 
                (mc:OfficeMonad<'c>) 
                (md:OfficeMonad<'d>) : OfficeMonad<'x> = 
        officeMonad { 
            let! a = ma
            let! b = mb
            let! c = mc
            let! d = md
            return (fn a b c d)
        }


    let liftM5 (fn:'a -> 'b -> 'c -> 'd -> 'e -> 'x) 
                (ma:OfficeMonad<'a>) 
                (mb:OfficeMonad<'b>) 
                (mc:OfficeMonad<'c>) 
                (md:OfficeMonad<'d>) 
                (me:OfficeMonad<'e>) : OfficeMonad<'x> = 
        officeMonad { 
            let! a = ma
            let! b = mb
            let! c = mc
            let! d = md
            let! e = me
            return (fn a b c d e)
        }

    let liftM6 (fn:'a -> 'b -> 'c -> 'd -> 'e -> 'f -> 'x) 
                (ma:OfficeMonad<'a>) 
                (mb:OfficeMonad<'b>) 
                (mc:OfficeMonad<'c>) 
                (md:OfficeMonad<'d>) 
                (me:OfficeMonad<'e>) 
                (mf:OfficeMonad<'f>) : OfficeMonad<'x> = 
        officeMonad { 
            let! a = ma
            let! b = mb
            let! c = mc
            let! d = md
            let! e = me
            let! f = mf
            return (fn a b c d e f)
        }


    let tupleM2 (ma:OfficeMonad<'a>) 
                (mb:OfficeMonad<'b>) : OfficeMonad<'a * 'b> = 
        liftM2 (fun a b -> (a,b)) ma mb

    let tupleM3 (ma:OfficeMonad<'a>) 
                (mb:OfficeMonad<'b>) 
                (mc:OfficeMonad<'c>) : OfficeMonad<'a * 'b * 'c> = 
        liftM3 (fun a b c -> (a,b,c)) ma mb mc

    let tupleM4 (ma:OfficeMonad<'a>) 
                (mb:OfficeMonad<'b>) 
                (mc:OfficeMonad<'c>) 
                (md:OfficeMonad<'d>) : OfficeMonad<'a * 'b * 'c * 'd> = 
        liftM4 (fun a b c d -> (a,b,c,d)) ma mb mc md

    let tupleM5 (ma:OfficeMonad<'a>) 
                (mb:OfficeMonad<'b>) 
                (mc:OfficeMonad<'c>) 
                (md:OfficeMonad<'d>) 
                (me:OfficeMonad<'e>) : OfficeMonad<'a * 'b * 'c * 'd * 'e> = 
        liftM5 (fun a b c d e -> (a,b,c,d,e)) ma mb mc md me

    let tupleM6 (ma:OfficeMonad<'a>) 
                (mb:OfficeMonad<'b>) 
                (mc:OfficeMonad<'c>) 
                (md:OfficeMonad<'d>) 
                (me:OfficeMonad<'e>) 
                (mf:OfficeMonad<'f>) : OfficeMonad<'a * 'b * 'c * 'd * 'e * 'f> = 
        liftM6 (fun a b c d e f -> (a,b,c,d,e,f)) ma mb mc md me mf


    let pipeM2 (ma:OfficeMonad<'a>) 
               (mb:OfficeMonad<'b>) 
               (fn:'a -> 'b -> 'x) : OfficeMonad<'x> = 
        liftM2 fn ma mb

    let pipeM3 (ma:OfficeMonad<'a>) 
               (mb:OfficeMonad<'b>) 
               (mc:OfficeMonad<'c>) 
               (fn:'a -> 'b -> 'c -> 'x): OfficeMonad<'x> = 
        liftM3 fn ma mb mc

    let pipeM4 (ma:OfficeMonad<'a>) 
               (mb:OfficeMonad<'b>) 
               (mc:OfficeMonad<'c>) 
               (md:OfficeMonad<'d>) 
               (fn:'a -> 'b -> 'c -> 'd -> 'x) : OfficeMonad<'x> = 
        liftM4 fn ma mb mc md

    let pipeM5 (ma:OfficeMonad<'a>) 
               (mb:OfficeMonad<'b>) 
               (mc:OfficeMonad<'c>) 
               (md:OfficeMonad<'d>) 
               (me:OfficeMonad<'e>) 
               (fn:'a -> 'b -> 'c -> 'd -> 'e ->'x) : OfficeMonad<'x> = 
        liftM5 fn ma mb mc md me

    let pipeM6 (ma:OfficeMonad<'a>) 
               (mb:OfficeMonad<'b>) 
               (mc:OfficeMonad<'c>) 
               (md:OfficeMonad<'d>) 
               (me:OfficeMonad<'e>) 
               (mf:OfficeMonad<'f>) 
               (fn:'a -> 'b -> 'c -> 'd -> 'e -> 'f -> 'x) : OfficeMonad<'x> = 
        liftM6 fn ma mb mc md me mf

    /// Left biased choice, if ``ma`` succeeds return its 
    /// result, otherwise try ``mb``.
    let altM (ma:OfficeMonad<'a>) (mb:OfficeMonad<'a>) : OfficeMonad<'a> = 
        combineM ma mb


    /// Haskell Applicative's (<*>)
    let apM (mf:OfficeMonad<'a ->'b>) 
            (ma:OfficeMonad<'a>) : OfficeMonad<'b> = 
        officeMonad { 
            let! fn = mf
            let! a = ma
            return (fn a) 
        }

    /// Perform two actions in sequence. 
    /// Ignore the results of the second action if both succeed.
    let seqL (ma:OfficeMonad<'a>) (mb:OfficeMonad<'b>) : OfficeMonad<'a> = 
        officeMonad { 
            let! a = ma
            let! b = mb
            return a
        }

    /// Perform two actions in sequence. 
    /// Ignore the results of the first action if both succeed.
    let seqR (ma:OfficeMonad<'a>) (mb:OfficeMonad<'b>) : OfficeMonad<'b> = 
        officeMonad { 
            let! a = ma
            let! b = mb
            return b
        }

    /// Optionally run a computation. 
    /// If the build fails return None otherwise retun Some<'a>.
    let optional (ma:OfficeMonad<'a>) : OfficeMonad<'a option> = 
        OfficeMonad <| fun env ->
            DocMonad.optional (apply1 ma env)

    let kleisliL (mf:'a -> OfficeMonad<'b>)
                 (mg:'b -> OfficeMonad<'c>)
                 (source:'a) : OfficeMonad<'c> = 
        officeMonad { 
            let! b = mf source
            let! c = mg b
            return c
        }

    let kleisliR (mf:'b -> OfficeMonad<'c>)
                 (mg:'a -> OfficeMonad<'b>)
                 (source:'a) : OfficeMonad<'c> = 
        officeMonad { 
            let! b = mg source
            let! c = mf b
            return c
        }

    // ****************************************************
    // Use Office applications...

    let execWord (operation:Word.Application -> BuildResult<'a>) : OfficeMonad<'a> =
        OfficeMonad <| fun env -> 
            try 
                match operation env.WordApp with
                | Ok a -> DocMonad.breturn a
                | Error msg -> DocMonad.throwError msg
            with
            | _ -> DocMonad.throwError "Resource error - Word.Application"


    let execExcel (operation:Excel.Application -> BuildResult<'a>) : OfficeMonad<'a> =
        OfficeMonad <| fun env -> 
            try 
                match operation env.ExcelApp with
                | Ok a -> DocMonad.breturn a
                | Error msg -> DocMonad.throwError msg
            with
            | _ -> DocMonad.throwError "Resource error - Excel.Application"

    let execPowerPoint (operation:PowerPoint.Application -> BuildResult<'a>) : OfficeMonad<'a> =
        OfficeMonad <| fun env -> 
            try 
                match operation env.PowerPointApp with
                | Ok a -> DocMonad.breturn a
                | Error msg -> DocMonad.throwError msg
            with
            | _ -> DocMonad.throwError "Resource error - PowerPoint.Application"

    // ****************************************************
    // Recursive functions


    /// Implementation wraps DocMonad.mapM which is in CPS
    let mapM (mf: 'a -> OfficeMonad<'b>) 
             (source:'a list) : OfficeMonad<'b list> = 
        OfficeMonad <| fun env -> 
            DocMonad.mapM (fun a -> apply1 (mf a) env) source


    /// Flipped mapM
    let forM (source:'a list) 
             (mf: 'a -> OfficeMonad<'b>) : OfficeMonad<'b list> = 
        mapM mf source

    /// Forgetful mapM
    let mapMz (mf: 'a -> OfficeMonad<'b>) 
              (source:'a list) : OfficeMonad<unit> = 
        OfficeMonad <| fun env -> 
            DocMonad.mapMz (fun a -> apply1 (mf a) env) source

    /// Flipped mapMz
    let forMz (source:'a list) (mf: 'a -> OfficeMonad<'b>) : OfficeMonad<unit> = 
        mapMz mf source

    /// Implemented in CPS 
    let mapiM (mf: int -> 'a -> OfficeMonad<'b>) 
              (source:'a list) : OfficeMonad<'b list> = 
        OfficeMonad <| fun env -> 
            DocMonad.mapiM (fun ix a -> apply1 (mf ix a) env) source

    /// Flipped mapMi
    let foriM (source:'a list) 
              (mf: int -> 'a -> OfficeMonad<'b>) : OfficeMonad<'b list> = 
        mapiM mf source

    /// Forgetful mapiM
    let mapiMz (mf: int -> 'a -> OfficeMonad<'b>) 
               (source:'a list) : OfficeMonad<unit> =
        OfficeMonad <| fun env -> 
            DocMonad.mapiMz (fun ix a -> apply1 (mf ix a) env) source

    /// Flipped mapiMz
    let foriMz (source:'a list) 
               (mf: int -> 'a -> OfficeMonad<'b>) : OfficeMonad<unit> = 
        mapiMz mf source


