﻿// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Base


module DocMonad = 

    open DocBuild.Base
    open DocBuild.Base.Shell

    type BuilderEnv = 
        { WorkingDirectory: string
          GhostscriptExe: string
          PdftkExe: string 
          PandocExe: string
          PandocReferenceDoc: string option
        }



    type DocMonad<'a> = 
        DocMonad of (BuilderEnv -> BuildResult<'a>)

    let inline private apply1 (ma: DocMonad<'a>) 
                                (env: BuilderEnv) : BuildResult<'a>= 
        let (DocMonad f) = ma in f env

    let inline breturn (x:'a) : DocMonad<'a> = 
        DocMonad <| fun _ -> Ok x

    let inline private bindM (ma:DocMonad<'a>) 
                        (f :'a -> DocMonad<'b>) : DocMonad<'b> =
        DocMonad <| fun env -> 
            match apply1 ma env with
            | Error msg -> Error msg
            | Ok a -> apply1 (f a) env

    let inline bzero () : DocMonad<'a> = 
        DocMonad <| fun _ -> Error "bzero"

    /// "First success"
    let inline private combineM (ma:DocMonad<'a>) 
                                    (mb:DocMonad<'a>) : DocMonad<'a> = 
        DocMonad <| fun env -> 
            match apply1 ma env with
            | Error msg -> apply1 mb env
            | Ok a -> Ok a


    let inline private  delayM (fn:unit -> DocMonad<'a>) : DocMonad<'a> = 
        bindM (breturn ()) fn 

    type DocMonadBuilder() = 
        member self.Return x            = breturn x
        member self.Bind (p,f)          = bindM p f
        member self.Zero ()             = bzero ()
        member self.Combine (ma,mb)     = combineM ma mb
        member self.Delay fn            = delayM fn

    let (docMonad:DocMonadBuilder) = new DocMonadBuilder()


    // ****************************************************
    // Run

    let runDocMonad (config:BuilderEnv) 
                    (ma:DocMonad<'a>) : BuildResult<'a> = 
        apply1 ma config

    let execDocMonad (config:BuilderEnv) (ma:DocMonad<'a>) : 'a = 
        match apply1 ma config with
        | Ok a -> a
        | Error msg -> failwith msg




    // ****************************************************
    // Errors

    let throwError (msg:string) : DocMonad<'a> = 
        DocMonad <| fun _ -> Error msg

    let swapError (msg:string) (ma:DocMonad<'a>) : DocMonad<'a> = 
        DocMonad <| fun env ->
            match apply1 ma env with
            | Error _ -> Error msg
            | Ok a -> Ok a

    let (<&?>) (ma:DocMonad<'a>) (msg:string) : DocMonad<'a> = 
        swapError msg ma

    let (<?&>) (msg:string) (ma:DocMonad<'a>) : DocMonad<'a> = 
        swapError msg ma

    // ****************************************************
    // Reader

    let ask : DocMonad<BuilderEnv> = 
        DocMonad <| fun env -> Ok env

    let asks (extract:BuilderEnv -> 'a) : DocMonad<'a> = 
        DocMonad <| fun env -> Ok (extract env)

    let local (update:BuilderEnv -> BuilderEnv) 
                (ma:DocMonad<'a>) : DocMonad<'a> = 
        DocMonad <| fun env -> apply1 ma (update env)

    // ****************************************************
    // Lift operations

    let liftDocMonad (answer:BuildResult<'a>) : DocMonad<'a> = 
        DocMonad <| fun _ -> answer

 

    // ****************************************************
    // Monadic operations

    /// Bind operator
    let (>>=) (ma:DocMonad<'a>) 
                (fn:'a -> DocMonad<'b>) : DocMonad<'b> = 
        bindM ma fn

    // Common monadic operations
    let fmapM (fn:'a -> 'b) (ma:DocMonad<'a>) : DocMonad<'b> = 
        DocMonad <| fun env -> 
           match apply1 ma env with
           | Error msg -> Error msg
           | Ok a -> Ok (fn a)

    /// Operator for fmap.
    let (|>>) (ma:DocMonad<'a>) (fn:'a -> 'b) : DocMonad<'b> = 
        fmapM fn ma

    /// Flipped fmap.
    let (<<|) (fn:'a -> 'b) (ma:DocMonad<'a>) : DocMonad<'b> = 
        fmapM fn ma

    // liftM (which is fmap)
    let liftM (fn:'a -> 'x) (ma:DocMonad<'a>) : DocMonad<'x> = 
        fmapM fn ma

    let liftM2 (fn:'a -> 'b -> 'x) 
                (ma:DocMonad<'a>) 
                (mb:DocMonad<'b>) : DocMonad<'x> = 
        docMonad { 
            let! a = ma
            let! b = mb
            return (fn a b)
        }

    let liftM3 (fn:'a -> 'b -> 'c -> 'x) 
                (ma:DocMonad<'a>) 
                (mb:DocMonad<'b>) 
                (mc:DocMonad<'c>) : DocMonad<'x> = 
        docMonad { 
            let! a = ma
            let! b = mb
            let! c = mc
            return (fn a b c)
        }

    let liftM4 (fn:'a -> 'b -> 'c -> 'd -> 'x) 
                (ma:DocMonad<'a>) 
                (mb:DocMonad<'b>) 
                (mc:DocMonad<'c>) 
                (md:DocMonad<'d>) : DocMonad<'x> = 
        docMonad { 
            let! a = ma
            let! b = mb
            let! c = mc
            let! d = md
            return (fn a b c d)
        }


    let liftM5 (fn:'a -> 'b -> 'c -> 'd -> 'e -> 'x) 
                (ma:DocMonad<'a>) 
                (mb:DocMonad<'b>) 
                (mc:DocMonad<'c>) 
                (md:DocMonad<'d>) 
                (me:DocMonad<'e>) : DocMonad<'x> = 
        docMonad { 
            let! a = ma
            let! b = mb
            let! c = mc
            let! d = md
            let! e = me
            return (fn a b c d e)
        }

    let liftM6 (fn:'a -> 'b -> 'c -> 'd -> 'e -> 'f -> 'x) 
                (ma:DocMonad<'a>) 
                (mb:DocMonad<'b>) 
                (mc:DocMonad<'c>) 
                (md:DocMonad<'d>) 
                (me:DocMonad<'e>) 
                (mf:DocMonad<'f>) : DocMonad<'x> = 
        docMonad { 
            let! a = ma
            let! b = mb
            let! c = mc
            let! d = md
            let! e = me
            let! f = mf
            return (fn a b c d e f)
        }


    let tupleM2 (ma:DocMonad<'a>) 
                (mb:DocMonad<'b>) : DocMonad<'a * 'b> = 
        liftM2 (fun a b -> (a,b)) ma mb

    let tupleM3 (ma:DocMonad<'a>) 
                (mb:DocMonad<'b>) 
                (mc:DocMonad<'c>) : DocMonad<'a * 'b * 'c> = 
        liftM3 (fun a b c -> (a,b,c)) ma mb mc

    let tupleM4 (ma:DocMonad<'a>) 
                (mb:DocMonad<'b>) 
                (mc:DocMonad<'c>) 
                (md:DocMonad<'d>) : DocMonad<'a * 'b * 'c * 'd> = 
        liftM4 (fun a b c d -> (a,b,c,d)) ma mb mc md

    let tupleM5 (ma:DocMonad<'a>) 
                (mb:DocMonad<'b>) 
                (mc:DocMonad<'c>) 
                (md:DocMonad<'d>) 
                (me:DocMonad<'e>) : DocMonad<'a * 'b * 'c * 'd * 'e> = 
        liftM5 (fun a b c d e -> (a,b,c,d,e)) ma mb mc md me

    let tupleM6 (ma:DocMonad<'a>) 
                (mb:DocMonad<'b>) 
                (mc:DocMonad<'c>) 
                (md:DocMonad<'d>) 
                (me:DocMonad<'e>) 
                (mf:DocMonad<'f>) : DocMonad<'a * 'b * 'c * 'd * 'e * 'f> = 
        liftM6 (fun a b c d e f -> (a,b,c,d,e,f)) ma mb mc md me mf

    let pipeM2 (ma:DocMonad<'a>) 
                (mb:DocMonad<'b>) 
                (fn:'a -> 'b -> 'x) : DocMonad<'x> = 
        liftM2 fn ma mb

    let pipeM3 (ma:DocMonad<'a>) 
                (mb:DocMonad<'b>) 
                (mc:DocMonad<'c>) 
                (fn:'a -> 'b -> 'c -> 'x): DocMonad<'x> = 
        liftM3 fn ma mb mc

    let pipeM4 (ma:DocMonad<'a>) 
                (mb:DocMonad<'b>) 
                (mc:DocMonad<'c>) 
                (md:DocMonad<'d>) 
                (fn:'a -> 'b -> 'c -> 'd -> 'x) : DocMonad<'x> = 
        liftM4 fn ma mb mc md

    let pipeM5 (ma:DocMonad<'a>) 
                (mb:DocMonad<'b>) 
                (mc:DocMonad<'c>) 
                (md:DocMonad<'d>) 
                (me:DocMonad<'e>) 
                (fn:'a -> 'b -> 'c -> 'd -> 'e ->'x): DocMonad<'x> = 
        liftM5 fn ma mb mc md me

    let pipeM6 (ma:DocMonad<'a>) 
                (mb:DocMonad<'b>) 
                (mc:DocMonad<'c>) 
                (md:DocMonad<'d>) 
                (me:DocMonad<'e>) 
                (mf:DocMonad<'f>) 
                (fn:'a -> 'b -> 'c -> 'd -> 'e -> 'f -> 'x): DocMonad<'x> = 
        liftM6 fn ma mb mc md me mf

    /// Left biased choice, if ``ma`` succeeds return its result, otherwise try ``mb``.
    let altM (ma:DocMonad<'a>) (mb:DocMonad<'a>) : DocMonad<'a> = 
        combineM ma mb

    let (<||>) (ma:DocMonad<'a>) (mb:DocMonad<'a>) : DocMonad<'a> = 
        altM ma mb <&?> "(<||>)"

    /// Haskell Applicative's (<*>)
    let apM (mf:DocMonad<'a ->'b>) (ma:DocMonad<'a>) : DocMonad<'b> = 
        docMonad { 
            let! fn = mf
            let! a = ma
            return (fn a) 
        }

    /// Operator for apM
    let (<**>) (ma:DocMonad<'a -> 'b>) (mb:DocMonad<'a>) : DocMonad<'b> = 
        apM ma mb

    /// Operator for fmapM
    let (<&&>) (fn:'a -> 'b) (ma:DocMonad<'a>) : DocMonad<'b> = 
        fmapM fn ma


    /// Perform two actions in sequence. 
    /// Ignore the results of the second action if both succeed.
    let seqL (ma:DocMonad<'a>) (mb:DocMonad<'b>) : DocMonad<'a> = 
        docMonad { 
            let! a = ma
            let! b = mb
            return a
        }

    /// Perform two actions in sequence. 
    /// Ignore the results of the first action if both succeed.
    let seqR (ma:DocMonad<'a>) (mb:DocMonad<'b>) : DocMonad<'b> = 
        docMonad { 
            let! a = ma
            let! b = mb
            return b
        }

    /// Operator for seqL
    let (.>>) (ma:DocMonad<'a>) (mb:DocMonad<'b>) : DocMonad<'a> = 
        seqL ma mb

    /// Operator for seqR
    let (>>>.) (ma:DocMonad<'a>) (mb:DocMonad<'b>) : DocMonad<'b> = 
        seqR ma mb

    /// Optionally run a computation. 
    /// If the build fails return None otherwise retun Some<'a>.
    let optional (ma:DocMonad<'a>) : DocMonad<'a option> = 
        DocMonad <| fun env ->
            match apply1 ma env with
            | Error _ -> Ok None
            | Ok a -> Ok (Some a)

   // ****************************************************
    // Execute 'builtin' processes 
    // (Respective applications must be installed)

    let private getOptions (findExe:BuilderEnv -> string) : DocMonad<ProcessOptions> = 
        pipeM2 (asks findExe)
                (asks (fun env -> env.WorkingDirectory))
                (fun exe cwd -> { WorkingDirectory = cwd; ExecutableName = exe})
    
    let private shellExecute (findExe:BuilderEnv -> string)
                                 (command:CommandArgs) : DocMonad<string> = 
        docMonad { 
            let! options = getOptions findExe
            let! ans = liftDocMonad <| executeProcess options command.Command
            return ans
            }
        
    let execGhostscript (command:CommandArgs) : DocMonad<string> = 
        shellExecute (fun env -> env.GhostscriptExe) command

    let execPandoc (command:CommandArgs) : DocMonad<string> = 
        shellExecute (fun env -> env.PandocExe) command

    let execPdftk (command:CommandArgs) : DocMonad<string> = 
        shellExecute (fun env -> env.PdftkExe) command

