// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Office.Monad


[<AutoOpen>]
module Monad = 

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
        member x.WordApp() : Word.Application = 
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

    type DocOffice<'a> = 
        DocOffice of (OfficeHandles -> Monad.DocBuild<'a>)

    let inline private apply1 (ma: DocOffice<'a>) 
                                (env: OfficeHandles) : Monad.DocBuild<'a> = 
        let (DocOffice f) = ma in f env

    let inline oreturn (x:'a) : DocOffice<'a> = 
        DocOffice <| fun _ -> Monad.breturn x

    let inline bindM (ma:DocOffice<'a>) 
                        (f :'a -> DocOffice<'b>) : DocOffice<'b> =
        DocOffice <| fun env -> 
            Monad.docBuild.Bind(apply1 ma env, fun a -> apply1 (f a) env)

    let inline ozero () : DocOffice<'a> = 
        DocOffice <| fun _ -> Monad.bzero () 

    /// "First success"
    let inline private combineM (ma:DocOffice<'a>) 
                                    (mb:DocOffice<'a>) : DocOffice<'a> = 
        DocOffice <| fun env -> 
            Monad.docBuild.Combine(apply1 ma env, apply1 mb env)


    let inline private  delayM (fn:unit -> DocOffice<'a>) : DocOffice<'a> = 
        bindM (oreturn ()) fn 

    type DocOfficeBuilder() = 
        member self.Return x            = oreturn x
        member self.Bind (p,f)          = bindM p f
        member self.Zero ()             = ozero ()
        member self.Combine (ma,mb)     = combineM ma mb
        member self.Delay fn            = delayM fn

    let (docOffice:DocOfficeBuilder) = new DocOfficeBuilder()


    // ****************************************************
    // Run
    let runDocBuild (config:Monad.BuilderEnv) 
                    (ma:DocOffice<'a>) : BuildResult<'a> = 
        let handles = new OfficeHandles()
        let (DocOffice mf) = ma

        let ans = Monad.runDocBuild config (mf handles)
        handles.RunFinalizer ()
        ans


    let execDocBuild (config:Monad.BuilderEnv) 
                        (ma:DocOffice<'a>) : 'a = 
        match runDocBuild config ma with
        | Ok a -> a
        | Error msg -> failwith msg
