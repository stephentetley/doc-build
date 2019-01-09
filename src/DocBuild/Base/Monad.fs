// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Base


module Monad = 

    open DocBuild.Base
    open DocBuild.Base.Shell

    type BuilderEnv = 
        { WorkingDirectory: string
          GhostscriptExe: string
          PdftkExe: string 
          PandocExe: string
          PandocReferenceDoc: string option
        }



    type DocBuild<'a> = 
        DocBuild of (BuilderEnv -> BuildResult<'a>)

    let inline private apply1 (ma: DocBuild<'a>) 
                                (env: BuilderEnv) : BuildResult<'a>= 
        let (DocBuild f) = ma in f env

    let inline breturn (x:'a) : DocBuild<'a> = 
        DocBuild <| fun _ -> Ok x

    let inline private bindM (ma:DocBuild<'a>) 
                        (f :'a -> DocBuild<'b>) : DocBuild<'b> =
        DocBuild <| fun env -> 
            match apply1 ma env with
            | Error msg -> Error msg
            | Ok a -> apply1 (f a) env

    let inline bzero () : DocBuild<'a> = 
        DocBuild <| fun _ -> Error "bzero"

    /// "First success"
    let inline private combineM (ma:DocBuild<'a>) 
                                    (mb:DocBuild<'a>) : DocBuild<'a> = 
        DocBuild <| fun env -> 
            match apply1 ma env with
            | Error msg -> apply1 mb env
            | Ok a -> Ok a


    let inline private  delayM (fn:unit -> DocBuild<'a>) : DocBuild<'a> = 
        bindM (breturn ()) fn 

    type DocBuildBuilder() = 
        member self.Return x            = breturn x
        member self.Bind (p,f)          = bindM p f
        member self.Zero ()             = bzero ()
        member self.Combine (ma,mb)     = combineM ma mb
        member self.Delay fn            = delayM fn

    let (docBuild:DocBuildBuilder) = new DocBuildBuilder()


    // ****************************************************
    // Run

    let runDocBuild (config:BuilderEnv) 
                    (ma:DocBuild<'a>) : BuildResult<'a> = 
        apply1 ma config

    let execDocBuild (config:BuilderEnv) (ma:DocBuild<'a>) : 'a = 
        match apply1 ma config with
        | Ok a -> a
        | Error msg -> failwith msg




    // ****************************************************
    // Errors

    let throwError (msg:string) : DocBuild<'a> = 
        DocBuild <| fun _ -> Error msg

    let swapError (msg:string) (ma:DocBuild<'a>) : DocBuild<'a> = 
        DocBuild <| fun env ->
            match apply1 ma env with
            | Error _ -> Error msg
            | Ok a -> Ok a

    let (<&?>) (ma:DocBuild<'a>) (msg:string) : DocBuild<'a> = 
        swapError msg ma

    let (<?&>) (msg:string) (ma:DocBuild<'a>) : DocBuild<'a> = 
        swapError msg ma

    // ****************************************************
    // Reader

    let ask : DocBuild<BuilderEnv> = 
        DocBuild <| fun env -> Ok env

    let asks (extract:BuilderEnv -> 'a) : DocBuild<'a> = 
        DocBuild <| fun env -> Ok (extract env)

    let local (update:BuilderEnv -> BuilderEnv) 
                (ma:DocBuild<'a>) : DocBuild<'a> = 
        DocBuild <| fun env -> apply1 ma (update env)

    // ****************************************************
    // Lift operations

    let liftDocBuild (answer:BuildResult<'a>) : DocBuild<'a> = 
        DocBuild <| fun _ -> answer

 

    // ****************************************************
    // Monadic operations

    /// Bind operator
    let (>>=) (ma:DocBuild<'a>) 
                (fn:'a -> DocBuild<'b>) : DocBuild<'b> = 
        bindM ma fn

    // Common monadic operations
    let fmapM (fn:'a -> 'b) (ma:DocBuild<'a>) : DocBuild<'b> = 
        DocBuild <| fun env -> 
           match apply1 ma env with
           | Error msg -> Error msg
           | Ok a -> Ok (fn a)

    /// Operator for fmap.
    let (|>>) (ma:DocBuild<'a>) (fn:'a -> 'b) : DocBuild<'b> = 
        fmapM fn ma

    /// Flipped fmap.
    let (<<|) (fn:'a -> 'b) (ma:DocBuild<'a>) : DocBuild<'b> = 
        fmapM fn ma

    // liftM (which is fmap)
    let liftM (fn:'a -> 'x) (ma:DocBuild<'a>) : DocBuild<'x> = 
        fmapM fn ma

    let liftM2 (fn:'a -> 'b -> 'x) 
                (ma:DocBuild<'a>) 
                (mb:DocBuild<'b>) : DocBuild<'x> = 
        docBuild { 
            let! a = ma
            let! b = mb
            return (fn a b)
        }

    let liftM3 (fn:'a -> 'b -> 'c -> 'x) 
                (ma:DocBuild<'a>) 
                (mb:DocBuild<'b>) 
                (mc:DocBuild<'c>) : DocBuild<'x> = 
        docBuild { 
            let! a = ma
            let! b = mb
            let! c = mc
            return (fn a b c)
        }

    let liftM4 (fn:'a -> 'b -> 'c -> 'd -> 'x) 
                (ma:DocBuild<'a>) 
                (mb:DocBuild<'b>) 
                (mc:DocBuild<'c>) 
                (md:DocBuild<'d>) : DocBuild<'x> = 
        docBuild { 
            let! a = ma
            let! b = mb
            let! c = mc
            let! d = md
            return (fn a b c d)
        }


    let liftM5 (fn:'a -> 'b -> 'c -> 'd -> 'e -> 'x) 
                (ma:DocBuild<'a>) 
                (mb:DocBuild<'b>) 
                (mc:DocBuild<'c>) 
                (md:DocBuild<'d>) 
                (me:DocBuild<'e>) : DocBuild<'x> = 
        docBuild { 
            let! a = ma
            let! b = mb
            let! c = mc
            let! d = md
            let! e = me
            return (fn a b c d e)
        }

    let liftM6 (fn:'a -> 'b -> 'c -> 'd -> 'e -> 'f -> 'x) 
                (ma:DocBuild<'a>) 
                (mb:DocBuild<'b>) 
                (mc:DocBuild<'c>) 
                (md:DocBuild<'d>) 
                (me:DocBuild<'e>) 
                (mf:DocBuild<'f>) : DocBuild<'x> = 
        docBuild { 
            let! a = ma
            let! b = mb
            let! c = mc
            let! d = md
            let! e = me
            let! f = mf
            return (fn a b c d e f)
        }


    let tupleM2 (ma:DocBuild<'a>) 
                (mb:DocBuild<'b>) : DocBuild<'a * 'b> = 
        liftM2 (fun a b -> (a,b)) ma mb

    let tupleM3 (ma:DocBuild<'a>) 
                (mb:DocBuild<'b>) 
                (mc:DocBuild<'c>) : DocBuild<'a * 'b * 'c> = 
        liftM3 (fun a b c -> (a,b,c)) ma mb mc

    let tupleM4 (ma:DocBuild<'a>) 
                (mb:DocBuild<'b>) 
                (mc:DocBuild<'c>) 
                (md:DocBuild<'d>) : DocBuild<'a * 'b * 'c * 'd> = 
        liftM4 (fun a b c d -> (a,b,c,d)) ma mb mc md

    let tupleM5 (ma:DocBuild<'a>) 
                (mb:DocBuild<'b>) 
                (mc:DocBuild<'c>) 
                (md:DocBuild<'d>) 
                (me:DocBuild<'e>) : DocBuild<'a * 'b * 'c * 'd * 'e> = 
        liftM5 (fun a b c d e -> (a,b,c,d,e)) ma mb mc md me

    let tupleM6 (ma:DocBuild<'a>) 
                (mb:DocBuild<'b>) 
                (mc:DocBuild<'c>) 
                (md:DocBuild<'d>) 
                (me:DocBuild<'e>) 
                (mf:DocBuild<'f>) : DocBuild<'a * 'b * 'c * 'd * 'e * 'f> = 
        liftM6 (fun a b c d e f -> (a,b,c,d,e,f)) ma mb mc md me mf

    let pipeM2 (ma:DocBuild<'a>) 
                (mb:DocBuild<'b>) 
                (fn:'a -> 'b -> 'x) : DocBuild<'x> = 
        liftM2 fn ma mb

    let pipeM3 (ma:DocBuild<'a>) 
                (mb:DocBuild<'b>) 
                (mc:DocBuild<'c>) 
                (fn:'a -> 'b -> 'c -> 'x): DocBuild<'x> = 
        liftM3 fn ma mb mc

    let pipeM4 (ma:DocBuild<'a>) 
                (mb:DocBuild<'b>) 
                (mc:DocBuild<'c>) 
                (md:DocBuild<'d>) 
                (fn:'a -> 'b -> 'c -> 'd -> 'x) : DocBuild<'x> = 
        liftM4 fn ma mb mc md

    let pipeM5 (ma:DocBuild<'a>) 
                (mb:DocBuild<'b>) 
                (mc:DocBuild<'c>) 
                (md:DocBuild<'d>) 
                (me:DocBuild<'e>) 
                (fn:'a -> 'b -> 'c -> 'd -> 'e ->'x): DocBuild<'x> = 
        liftM5 fn ma mb mc md me

    let pipeM6 (ma:DocBuild<'a>) 
                (mb:DocBuild<'b>) 
                (mc:DocBuild<'c>) 
                (md:DocBuild<'d>) 
                (me:DocBuild<'e>) 
                (mf:DocBuild<'f>) 
                (fn:'a -> 'b -> 'c -> 'd -> 'e -> 'f -> 'x): DocBuild<'x> = 
        liftM6 fn ma mb mc md me mf

    /// Left biased choice, if ``ma`` succeeds return its result, otherwise try ``mb``.
    let altM (ma:DocBuild<'a>) (mb:DocBuild<'a>) : DocBuild<'a> = 
        combineM ma mb

    let (<||>) (ma:DocBuild<'a>) (mb:DocBuild<'a>) : DocBuild<'a> = 
        altM ma mb <&?> "(<||>)"

    /// Haskell Applicative's (<*>)
    let apM (mf:DocBuild<'a ->'b>) (ma:DocBuild<'a>) : DocBuild<'b> = 
        docBuild { 
            let! fn = mf
            let! a = ma
            return (fn a) 
        }

    /// Operator for apM
    let (<**>) (ma:DocBuild<'a -> 'b>) (mb:DocBuild<'a>) : DocBuild<'b> = 
        apM ma mb

    /// Operator for fmapM
    let (<&&>) (fn:'a -> 'b) (ma:DocBuild<'a>) : DocBuild<'b> = 
        fmapM fn ma


    /// Perform two actions in sequence. 
    /// Ignore the results of the second action if both succeed.
    let seqL (ma:DocBuild<'a>) (mb:DocBuild<'b>) : DocBuild<'a> = 
        docBuild { 
            let! a = ma
            let! b = mb
            return a
        }

    /// Perform two actions in sequence. 
    /// Ignore the results of the first action if both succeed.
    let seqR (ma:DocBuild<'a>) (mb:DocBuild<'b>) : DocBuild<'b> = 
        docBuild { 
            let! a = ma
            let! b = mb
            return b
        }

    /// Operator for seqL
    let (.>>) (ma:DocBuild<'a>) (mb:DocBuild<'b>) : DocBuild<'a> = 
        seqL ma mb

    /// Operator for seqR
    let (>>>.) (ma:DocBuild<'a>) (mb:DocBuild<'b>) : DocBuild<'b> = 
        seqR ma mb

    /// Optionally run a computation. 
    /// If the build fails return None otherwise retun Some<'a>.
    let optional (ma:DocBuild<'a>) : DocBuild<'a option> = 
        DocBuild <| fun env ->
            match apply1 ma env with
            | Error _ -> Ok None
            | Ok a -> Ok (Some a)

   // ****************************************************
    // Execute 'builtin' processes 
    // (Respective applications must be installed)

    let private getOptions (findExe:BuilderEnv -> string) : DocBuild<ProcessOptions> = 
        pipeM2 (asks findExe)
                (asks (fun env -> env.WorkingDirectory))
                (fun exe cwd -> { WorkingDirectory = cwd; ExecutableName = exe})
    
    let private shellExecute (findExe:BuilderEnv -> string)
                                 (command:CommandArgs) : DocBuild<string> = 
        docBuild { 
            let! options = getOptions findExe
            let! ans = liftDocBuild <| executeProcess options command.Command
            return ans
            }
        
    let execGhostscript (command:CommandArgs) : DocBuild<string> = 
        shellExecute (fun env -> env.GhostscriptExe) command

    let execPandoc (command:CommandArgs) : DocBuild<string> = 
        shellExecute (fun env -> env.PandocExe) command

    let execPdftk (command:CommandArgs) : DocBuild<string> = 
        shellExecute (fun env -> env.PdftkExe) command

