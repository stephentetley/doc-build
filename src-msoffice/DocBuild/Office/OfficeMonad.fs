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


