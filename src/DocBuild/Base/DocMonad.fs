// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Base

// Explicitly import this module:
// open DocBuild.Base.DocMonad

module DocMonad = 

    open System

    open DocBuild.Base
    open DocBuild.Base.Shell

    type BuilderEnv = 
        { WorkingDirectory: Uri
          SourceDirectory: Uri
          IncludeDirectory: Uri
          GhostscriptExe: string
          PdftkExe: string 
          PandocExe: string
        }



    /// DocMonad is parametric on 'res (user resources)
    /// This allows a level of extensibility on the applications
    /// that DocMonad can run (e.g. Office apps Word, Excel)

    type ResourceFinalize =
        abstract RunFinalizer : unit

    type DocMonad<'res,'a> = 
        DocMonad of ('res -> BuilderEnv -> BuildResult<'a>)

    let inline private apply1 (ma: DocMonad<'res,'a>) 
                              (res:'res) 
                              (env: BuilderEnv) : BuildResult<'a>= 
        let (DocMonad f) = ma in f res env

    let inline dreturn (x:'a) : DocMonad<'res,'a> = 
        DocMonad <| fun _ _ -> Ok x

    let inline private bindM (ma:DocMonad<'res,'a>) 
                        (f :'a -> DocMonad<'res,'b>) : DocMonad<'res,'b> =
        DocMonad <| fun res env -> 
            match apply1 ma res env with
            | Error msg -> Error msg
            | Ok a -> apply1 (f a) res env

    let inline dzero () : DocMonad<'res,'a> = 
        DocMonad <| fun _ _ -> Error "bzero"

    /// "First success"
    let inline private combineM (ma:DocMonad<'res,'a>) 
                                (mb:DocMonad<'res,'a>) : DocMonad<'res,'a> = 
        DocMonad <| fun res env -> 
            match apply1 ma res env with
            | Error msg -> apply1 mb res env
            | Ok a -> Ok a


    let inline private  delayM (fn:unit -> DocMonad<'res,'a>) : DocMonad<'res,'a> = 
        bindM (dreturn ()) fn 

    type DocMonadBuilder() = 
        member self.Return x            = dreturn x
        member self.Bind (p,f)          = bindM p f
        member self.Zero ()             = dzero ()
        member self.Combine (ma,mb)     = combineM ma mb
        member self.Delay fn            = delayM fn

    let (docMonad:DocMonadBuilder) = new DocMonadBuilder()


    // ****************************************************
    // Run

    /// This runs the finalizer on userResources
    let runDocMonad (userResources:#ResourceFinalize) 
                    (config:BuilderEnv) 
                    (ma:DocMonad<#ResourceFinalize,'a>) : BuildResult<'a> = 
        let ans = apply1 ma userResources config
        userResources.RunFinalizer |> ignore
        ans

    let runDocMonadNoCleanup (userResources:'res) 
                             (config:BuilderEnv) 
                             (ma:DocMonad<'res,'a>) : BuildResult<'a> = 
        apply1 ma userResources config
        
    let execDocMonad (userResources:#ResourceFinalize) 
                     (config:BuilderEnv) 
                     (ma:DocMonad<#ResourceFinalize,'a>) : 'a = 
        match runDocMonad userResources config ma with
        | Ok a -> a
        | Error msg -> failwith msg

    let execDocMonadNoCleanup (userResources:'res) 
                     (config:BuilderEnv) 
                     (ma:DocMonad<'res,'a>) : 'a = 
        match apply1 ma userResources config with
        | Ok a -> a
        | Error msg -> failwith msg




    // ****************************************************
    // Errors

    let throwError (msg:string) : DocMonad<'res,'a> = 
        DocMonad <| fun _ _ -> Error msg

    let swapError (msg:string) (ma:DocMonad<'res,'a>) : DocMonad<'res,'a> = 
        DocMonad <| fun res env ->
            match apply1 ma res env with
            | Error _ -> Error msg
            | Ok a -> Ok a


    /// Execute an action that may throw an exception.
    /// Capture the exception with try ... with
    /// and return the answer or the expection message in the monad.
    let attempt (ma: DocMonad<'res,'a>) : DocMonad<'res,'a> = 
        DocMonad <| fun res env -> 
            try
                apply1 ma res env
            with
            | ex -> Error (sprintf "attempt: %s" ex.Message)


    // ****************************************************
    // Reader

    let ask () : DocMonad<'res,BuilderEnv> = 
        DocMonad <| fun _ env -> Ok env

    let asks (extract:BuilderEnv -> 'a) : DocMonad<'res,'a> = 
        DocMonad <| fun _ env -> Ok (extract env)

    let askWorkingDirectory () : DocMonad<'res,Uri> = 
        asks (fun env -> env.WorkingDirectory)

    let askSourceDirectory () : DocMonad<'res,Uri> = 
        asks (fun env -> env.SourceDirectory)
        
    let askIncludeDirectory () : DocMonad<'res,Uri> = 
        asks (fun env -> env.IncludeDirectory)

    /// Use with caution.
    /// Generally you might only want to update the 
    /// working directory
    let local (update:BuilderEnv -> BuilderEnv) 
              (ma:DocMonad<'res,'a>) : DocMonad<'res,'a> = 
        DocMonad <| fun res env -> apply1 ma res (update env)



    let askUserResources () : DocMonad<'res,'res> = 
        DocMonad <| fun res _ -> Ok res

    let asksUserResources (extract:'res -> 'a) : DocMonad<'res,'a> = 
        DocMonad <| fun res _ -> Ok (extract res)

    /// Use with caution.
    /// Generally you might only want to update the 
    /// working directory
    let localUserResource (update:'res -> 'res) 
                          (ma:DocMonad<'res,'a>) : DocMonad<'res,'a> = 
        DocMonad <| fun res env -> apply1 ma (update res) env
            


    // ****************************************************
    // Lift operations

    let liftResult (answer:BuildResult<'a>) : DocMonad<'res,'a> = 
        DocMonad <| fun _ _ -> answer

 

    // ****************************************************
    // Monadic operations


    /// fmap 
    let fmapM (fn:'a -> 'b) (ma:DocMonad<'res,'a>) : DocMonad<'res,'b> = 
        DocMonad <| fun res env -> 
           match apply1 ma res env with
           | Error msg -> Error msg
           | Ok a -> Ok (fn a)


    // liftM (which is fmap)
    let liftM (fn:'a -> 'x) (ma:DocMonad<'res,'a>) : DocMonad<'res,'x> = 
        fmapM fn ma

    let liftM2 (fn:'a -> 'b -> 'x) 
               (ma:DocMonad<'res,'a>) 
               (mb:DocMonad<'res,'b>) : DocMonad<'res,'x> = 
        docMonad { 
            let! a = ma
            let! b = mb
            return (fn a b)
        }

    let liftM3 (fn:'a -> 'b -> 'c -> 'x) 
               (ma:DocMonad<'res,'a>) 
               (mb:DocMonad<'res,'b>) 
               (mc:DocMonad<'res,'c>) : DocMonad<'res,'x> = 
        docMonad { 
            let! a = ma
            let! b = mb
            let! c = mc
            return (fn a b c)
        }

    let liftM4 (fn:'a -> 'b -> 'c -> 'd -> 'x) 
               (ma:DocMonad<'res,'a>) 
               (mb:DocMonad<'res,'b>) 
               (mc:DocMonad<'res,'c>) 
               (md:DocMonad<'res,'d>) : DocMonad<'res,'x> = 
        docMonad { 
            let! a = ma
            let! b = mb
            let! c = mc
            let! d = md
            return (fn a b c d)
        }


    let liftM5 (fn:'a -> 'b -> 'c -> 'd -> 'e -> 'x) 
               (ma:DocMonad<'res,'a>) 
               (mb:DocMonad<'res,'b>) 
               (mc:DocMonad<'res,'c>) 
               (md:DocMonad<'res,'d>) 
               (me:DocMonad<'res,'e>) : DocMonad<'res,'x> = 
        docMonad { 
            let! a = ma
            let! b = mb
            let! c = mc
            let! d = md
            let! e = me
            return (fn a b c d e)
        }

    let liftM6 (fn:'a -> 'b -> 'c -> 'd -> 'e -> 'f -> 'x) 
               (ma:DocMonad<'res,'a>) 
               (mb:DocMonad<'res,'b>) 
               (mc:DocMonad<'res,'c>) 
               (md:DocMonad<'res,'d>) 
               (me:DocMonad<'res,'e>) 
               (mf:DocMonad<'res,'f>) : DocMonad<'res,'x> = 
        docMonad { 
            let! a = ma
            let! b = mb
            let! c = mc
            let! d = md
            let! e = me
            let! f = mf
            return (fn a b c d e f)
        }


    let tupleM2 (ma:DocMonad<'res,'a>) 
                (mb:DocMonad<'res,'b>) : DocMonad<'res,'a * 'b> = 
        liftM2 (fun a b -> (a,b)) ma mb

    let tupleM3 (ma:DocMonad<'res,'a>) 
                (mb:DocMonad<'res,'b>) 
                (mc:DocMonad<'res,'c>) : DocMonad<'res,'a * 'b * 'c> = 
        liftM3 (fun a b c -> (a,b,c)) ma mb mc

    let tupleM4 (ma:DocMonad<'res,'a>) 
                (mb:DocMonad<'res,'b>) 
                (mc:DocMonad<'res,'c>) 
                (md:DocMonad<'res,'d>) : DocMonad<'res,'a * 'b * 'c * 'd> = 
        liftM4 (fun a b c d -> (a,b,c,d)) ma mb mc md

    let tupleM5 (ma:DocMonad<'res,'a>) 
                (mb:DocMonad<'res,'b>) 
                (mc:DocMonad<'res,'c>) 
                (md:DocMonad<'res,'d>) 
                (me:DocMonad<'res,'e>) : DocMonad<'res,'a * 'b * 'c * 'd * 'e> = 
        liftM5 (fun a b c d e -> (a,b,c,d,e)) ma mb mc md me

    let tupleM6 (ma:DocMonad<'res,'a>) 
                (mb:DocMonad<'res,'b>) 
                (mc:DocMonad<'res,'c>) 
                (md:DocMonad<'res,'d>) 
                (me:DocMonad<'res,'e>) 
                (mf:DocMonad<'res,'f>) : DocMonad<'res,'a * 'b * 'c * 'd * 'e * 'f> = 
        liftM6 (fun a b c d e f -> (a,b,c,d,e,f)) ma mb mc md me mf

    let pipeM2 (ma:DocMonad<'res,'a>) 
               (mb:DocMonad<'res,'b>) 
               (fn:'a -> 'b -> 'x) : DocMonad<'res,'x> = 
        liftM2 fn ma mb

    let pipeM3 (ma:DocMonad<'res,'a>) 
               (mb:DocMonad<'res,'b>) 
               (mc:DocMonad<'res,'c>) 
               (fn:'a -> 'b -> 'c -> 'x) : DocMonad<'res,'x> = 
        liftM3 fn ma mb mc

    let pipeM4 (ma:DocMonad<'res,'a>) 
               (mb:DocMonad<'res,'b>) 
               (mc:DocMonad<'res,'c>) 
               (md:DocMonad<'res,'d>) 
               (fn:'a -> 'b -> 'c -> 'd -> 'x) : DocMonad<'res,'x> = 
        liftM4 fn ma mb mc md

    let pipeM5 (ma:DocMonad<'res,'a>) 
               (mb:DocMonad<'res,'b>) 
               (mc:DocMonad<'res,'c>) 
               (md:DocMonad<'res,'d>) 
               (me:DocMonad<'res,'e>) 
               (fn:'a -> 'b -> 'c -> 'd -> 'e ->'x) : DocMonad<'res,'x> = 
        liftM5 fn ma mb mc md me

    let pipeM6 (ma:DocMonad<'res,'a>) 
               (mb:DocMonad<'res,'b>) 
               (mc:DocMonad<'res,'c>) 
               (md:DocMonad<'res,'d>) 
               (me:DocMonad<'res,'e>) 
               (mf:DocMonad<'res,'f>) 
               (fn:'a -> 'b -> 'c -> 'd -> 'e -> 'f -> 'x) : DocMonad<'res,'x> = 
        liftM6 fn ma mb mc md me mf

    /// Left biased choice, if ``ma`` succeeds return its result, otherwise try ``mb``.
    let altM (ma:DocMonad<'res,'a>) (mb:DocMonad<'res,'a>) : DocMonad<'res,'a> = 
        combineM ma mb


    /// Haskell Applicative's (<*>)
    let apM (mf:DocMonad<'res,'a ->'b>) (ma:DocMonad<'res,'a>) : DocMonad<'res,'b> = 
        docMonad { 
            let! fn = mf
            let! a = ma
            return (fn a) 
        }



    /// Perform two actions in sequence. 
    /// Ignore the results of the second action if both succeed.
    let seqL (ma:DocMonad<'res,'a>) (mb:DocMonad<'res,'b>) : DocMonad<'res,'a> = 
        docMonad { 
            let! a = ma
            let! b = mb
            return a
        }

    /// Perform two actions in sequence. 
    /// Ignore the results of the first action if both succeed.
    let seqR (ma:DocMonad<'res,'a>) (mb:DocMonad<'res,'b>) : DocMonad<'res,'b> = 
        docMonad { 
            let! a = ma
            let! b = mb
            return b
        }


    /// Optionally run a computation. 
    /// If the build fails return None otherwise retun Some<'a>.
    let optional (ma:DocMonad<'res,'a>) : DocMonad<'res,'a option> = 
        DocMonad <| fun res env ->
            match apply1 ma res env with
            | Error _ -> Ok None
            | Ok a -> Ok (Some a)

    let kleisliL (mf:'a -> DocMonad<'res,'b>)
                 (mg:'b -> DocMonad<'res,'c>)
                 (source:'a) : DocMonad<'res,'c> = 
        docMonad { 
            let! b = mf source
            let! c = mg b
            return c
        }

    /// Flipped kleisliL
    let kleisliR (mf:'b -> DocMonad<'res,'c>)
                 (mg:'a -> DocMonad<'res,'b>)
                 (source:'a) : DocMonad<'res,'c> = 
        docMonad { 
            let! b = mg source
            let! c = mf b
            return c
        }


    // ****************************************************
    // Execute 'builtin' processes 
    // (Respective applications must be installed)

    let private getOptions (findExe:BuilderEnv -> string) : DocMonad<'res,ProcessOptions> = 
        pipeM2 (asks findExe)
                (asks (fun env -> env.WorkingDirectory))
                (fun exe cwd -> { WorkingDirectory = cwd.LocalPath
                                ; ExecutableName = exe})
    
    let private shellExecute (findExe:BuilderEnv -> string)
                                 (command:CommandArgs) : DocMonad<'res,string> = 
        docMonad { 
            let! options = getOptions findExe
            let! ans = liftResult <| executeProcess options command.Command
            return ans
            }
        
    let execGhostscript (command:CommandArgs) : DocMonad<'res,string> = 
        shellExecute (fun env -> env.GhostscriptExe) command

    let execPandoc (command:CommandArgs) : DocMonad<'res,string> = 
        shellExecute (fun env -> env.PandocExe) command

    let execPdftk (command:CommandArgs) : DocMonad<'res,string> = 
        shellExecute (fun env -> env.PdftkExe) command

    // ****************************************************
    // Recursive functions


    /// Implemented in CPS 
    let mapM (mf: 'a -> DocMonad<'res,'b>) 
             (source:'a list) : DocMonad<'res,'b list> = 
        DocMonad <| fun res env -> 
            let rec work ac ys fk sk = 
                match ys with
                | [] -> sk (List.rev ac)
                | z :: zs -> 
                    match apply1 (mf z) res env with
                    | Error msg -> fk msg
                    | Ok ans -> 
                        work ac zs fk (fun acs ->
                        sk (ans::acs))
            work [] source (fun msg -> Error msg) (fun ans -> Ok ans)

    /// Flipped mapM
    let forM (source:'a list) 
             (mf: 'a -> DocMonad<'res,'b>) : DocMonad<'res,'b list> = 
        mapM mf source

    /// Forgetful mapM
    let mapMz (mf: 'a -> DocMonad<'res,'b>) 
              (source:'a list) : DocMonad<'res,unit> = 
        DocMonad <| fun res env -> 
            let rec work ys cont = 
                match ys with
                | [] -> cont (Ok ())
                | z :: zs -> 
                    match apply1 (mf z) res env with
                    | Error msg -> cont (Error msg)
                    | Ok ans -> work zs cont
            work source id

    /// Flipped mapMz
    let forMz (source:'a list) 
              (mf: 'a -> DocMonad<'res,'b>) : DocMonad<'res,unit> = 
        mapMz mf source


    /// Implemented in CPS 
    let mapiM (mf:int -> 'a -> DocMonad<'res,'b>) 
              (source:'a list) : DocMonad<'res,'b list> = 
        DocMonad <| fun res env -> 
            let rec work ac n ys fk sk = 
                match ys with
                | [] -> sk (List.rev ac)
                | z :: zs -> 
                    match apply1 (mf n z) res env with
                    | Error msg -> fk msg
                    | Ok ans -> 
                        work ac (n+1) zs fk (fun acs ->
                        sk (ans::acs))
            work [] 0 source (fun msg -> Error msg) (fun ans -> Ok ans)

    /// Flipped mapMi
    let foriM (source:'a list) 
              (mf: int -> 'a -> DocMonad<'res,'b>)  : DocMonad<'res,'b list> = 
        mapiM mf source

    /// Forgetful mapiM
    let mapiMz (mf: int -> 'a -> DocMonad<'res,'b>) 
              (source:'a list) : DocMonad<'res,unit> = 
        DocMonad <| fun res env -> 
            let rec work n ys cont = 
                match ys with
                | [] -> cont (Ok ())
                | z :: zs -> 
                    match apply1 (mf n z) res env with
                    | Error msg -> cont (Error msg)
                    | Ok ans -> work (n+1) zs cont
            work 0 source id

    /// Flipped mapiMz
    let foriMz (source:'a list) 
               (mf: int -> 'a -> DocMonad<'res,'b>) : DocMonad<'res,unit> = 
        mapiMz mf source

  