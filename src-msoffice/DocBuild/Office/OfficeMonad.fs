// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Office.OfficeMonad


[<AutoOpen>]
module OfficeMonad = 

    open Microsoft.Office.Interop
    open DocBuild.Base

    open DocBuild.Office.Internal


    // Idea - maybe the 'res param in DocMonad<'res,'a>
    // should be an obj, this gives us an extensible type.
    // For instance, if we add add OfficeHandles a user might
    // still want to add a further handle to TeX / LaTeX.

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



    //// ****************************************************
    //// Run
    //let runOfficeMonad (config:DocMonad.BuilderEnv) 
    //                (ma:OfficeMonad<'a>) : BuildResult<'a> = 
    //    let handles = new OfficeHandles()
    //    let (OfficeMonad mf) = ma

    //    let ans = DocMonad.runDocMonad config (mf handles)
    //    handles.RunFinalizer ()
    //    ans


    //let execOfficeMonad (config:DocMonad.BuilderEnv) 
    //                    (ma:OfficeMonad<'a>) : 'a = 
    //    match runOfficeMonad config ma with
    //    | Ok a -> a
    //    | Error msg -> failwith msg



    //// ****************************************************
    //// Use Office applications...

    //let execWord (operation:Word.Application -> BuildResult<'a>) : OfficeMonad<'a> =
    //    OfficeMonad <| fun env -> 
    //        try 
    //            match operation env.WordApp with
    //            | Ok a -> DocMonad.breturn a
    //            | Error msg -> DocMonad.throwError msg
    //        with
    //        | _ -> DocMonad.throwError "Resource error - Word.Application"


    //let execExcel (operation:Excel.Application -> BuildResult<'a>) : OfficeMonad<'a> =
    //    OfficeMonad <| fun env -> 
    //        try 
    //            match operation env.ExcelApp with
    //            | Ok a -> DocMonad.breturn a
    //            | Error msg -> DocMonad.throwError msg
    //        with
    //        | _ -> DocMonad.throwError "Resource error - Excel.Application"

    //let execPowerPoint (operation:PowerPoint.Application -> BuildResult<'a>) : OfficeMonad<'a> =
    //    OfficeMonad <| fun env -> 
    //        try 
    //            match operation env.PowerPointApp with
    //            | Ok a -> DocMonad.breturn a
    //            | Error msg -> DocMonad.throwError msg
    //        with
    //        | _ -> DocMonad.throwError "Resource error - PowerPoint.Application"

