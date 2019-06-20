// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Base

// Explicitly import this module:
// open DocBuild.Base.DocMonad

[<AutoOpen>]
module DocMonad = 

    open System.IO

    open SLFormat.CommandOptions

    open DocBuild.Base


    /// If CustomStylesDocx is not an abs path, it will be looked for
    /// in the include directory.
    type PandocOptions = 
        { CustomStylesDocx: string option   
          PdfEngine: string option        /// usually "pdflatex"
        }

    type AppResources<'userRes> = 
        { PandocExe: string 
          GhostscriptExe: string
          PdftkExe: string 
          UserResources: 'userRes
        }

    /// PandocPdfEngine needs TeX installed
    type DocBuildEnv = 
        { WorkingDirectory: string
          SourceDirectory: string
          IncludeDirectories: string list          
          PandocOpts: PandocOptions
          PrintOrScreen: PrintQuality                
        }


    /// DocMonad is parametric on 'userRes (user resources)
    /// This allows a level of extensibility on the applications
    /// that DocMonad can run (e.g. Office apps Word, Excel)

    type IResourceFinalize =
        abstract RunFinalizer : unit

    type DocMonad<'a, 'userRes> = 
        DocMonad of (StreamWriter -> AppResources<'userRes> -> DocBuildEnv -> BuildResult<'a>)

    let inline private apply1 (ma: DocMonad<'a, 'userRes>) 
                              (sw: StreamWriter)
                              (res: AppResources<'userRes>) 
                              (env: DocBuildEnv) : BuildResult<'a>= 
        let (DocMonad f) = ma in f sw res env

    let inline mreturn (x:'a) : DocMonad<'a, 'userRes> = 
        DocMonad <| fun _ _ _ -> Ok x

    let inline private bindM (ma:DocMonad<'a, 'userRes>) 
                             (f :'a -> DocMonad<'b, 'userRes>) : DocMonad<'b, 'userRes> =
        DocMonad <| fun sw res env -> 
            match apply1 ma sw res env with
            | Error msg -> Error msg
            | Ok a -> apply1 (f a) sw res env

    let inline private zeroM () : DocMonad<'a, 'userRes> = 
        DocMonad <| fun _ _ _ -> Error "zeroM"

    /// "First success"
    let inline private combineM (ma:DocMonad<'a, 'userRes>) 
                                (mb:DocMonad<'a, 'userRes>) : DocMonad<'a, 'userRes> = 
        DocMonad <| fun sw res env -> 
            match apply1 ma sw res env with
            | Error msg -> apply1 mb sw res env
            | Ok a -> Ok a


    let inline private delayM (fn:unit -> DocMonad<'a, 'userRes>) : DocMonad<'a, 'userRes> = 
        bindM (mreturn ()) fn 

    type DocMonadBuilder() = 
        member self.Return x            = mreturn x
        member self.Bind (p,f)          = bindM p f
        member self.Zero ()             = zeroM ()
        member self.Combine (ma,mb)     = combineM ma mb
        member self.Delay fn            = delayM fn
        member self.ReturnFrom(ma)      = ma


    let (docMonad:DocMonadBuilder) = new DocMonadBuilder()


    // ****************************************************
    // Run

    /// This runs the finalizer on userResources
    let runDocMonad (resources:AppResources<#IResourceFinalize>) 
                    (config:DocBuildEnv) 
                    (ma:DocMonad<'a, #IResourceFinalize>) : BuildResult<'a> = 
        let logPath = Path.Combine (config.WorkingDirectory, "doc-build.log")
        use sw = new StreamWriter(path = logPath)
        let ans = apply1 ma sw resources config
        resources.UserResources.RunFinalizer |> ignore
        ans

    let runDocMonadNoCleanup (resources:AppResources<'userRes>) 
                             (config:DocBuildEnv) 
                             (ma:DocMonad<'a, 'userRes>) : BuildResult<'a> = 
        let logPath = Path.Combine (config.WorkingDirectory, "doc-build.log")
        use sw = new StreamWriter(path = logPath)
        apply1 ma sw resources config
        
    let execDocMonad (resources:AppResources<#IResourceFinalize>) 
                     (config:DocBuildEnv) 
                     (ma:DocMonad<'a, #IResourceFinalize>) : 'a = 
        match runDocMonad resources config ma with
        | Ok a -> a
        | Error msg -> failwith msg

    let execDocMonadNoCleanup (resources:AppResources<'userRes>) 
                              (config:DocBuildEnv) 
                              (ma:DocMonad<'a, 'userRes>) : 'a = 
        match runDocMonadNoCleanup resources config ma with
        | Ok a -> a
        | Error msg -> failwith msg


    /// fmap 
    let fmapM (fn:'a -> 'b) (ma:DocMonad<'a, 'userRes>) : DocMonad<'b, 'userRes> = 
        DocMonad <| fun sw res env -> 
           match apply1 ma sw res env with
           | Error msg -> Error msg
           | Ok a -> Ok (fn a)


    // ****************************************************
    // Errors

    let docError (msg:string) : DocMonad<'a, 'userRes> = 
        DocMonad <| fun _ _ _ -> Error msg

    let swapError (msg:string) (ma:DocMonad<'a, 'userRes>) : DocMonad<'a, 'userRes> = 
        DocMonad <| fun sw res env ->
            match apply1 ma sw res env with
            | Error _ -> Error msg
            | Ok a -> Ok a


    let ( <?> ) (ma:DocMonad<'a, 'userRes>) (msg:string) : DocMonad<'a, 'userRes> = 
        swapError msg ma

 

    let liftAssert (failMsg:string) (condition:bool) : DocMonad<unit, 'userRes> = 
        if condition then mreturn () else docError failMsg

    let liftOption (failMsg:string) (opt:'a option) : DocMonad<'a, 'userRes> = 
        match opt with
        | Some a -> mreturn a 
        | None -> docError failMsg




    let liftOperation (failMsg:string) 
                      (operation: unit -> 'a) : DocMonad<'a, 'userRes> = 
            try
                operation () |> mreturn
            with
            | ex -> docError failMsg


    let liftOperationResult (failMsg:string) (operation: unit -> Result<'a, string>) : DocMonad<'a, 'userRes> = 
        docMonad { 
            match! liftOperation failMsg operation with
            | Ok a -> return a
            | Error msg -> return! docError msg
        }


    let assertM (failMsg:string) (ma:DocMonad<bool, 'userRes>) : DocMonad<unit, 'userRes> = 
        bindM ma <| fun condition -> 
            if condition then 
                mreturn () 
            else docError failMsg


    let assertEmpty (source : 'a list) : DocMonad<unit, 'userRes> = 
        match source with
        | [] -> mreturn ()
        | _ -> docError "assertEmpty - non empty"
       

    let assertNonEmpty (source : 'a list) : DocMonad<'a list, 'userRes> = 
        match source with
        | [] -> docError "assertNonEmpty - empty"
        | _ -> mreturn source


    let assertSingleton (source : 'a list) : DocMonad<'a, 'userRes> = 
        match source with
        | [a] -> mreturn a
        | [] -> docError "assertSingleton - empty"
        | _ -> docError "assertSingleton - more than 1 item"


    let optionM (defaultValue:'a) 
                (ma:DocMonad<'a, 'userRes>) : DocMonad<'a, 'userRes> = 
        combineM ma (mreturn defaultValue)

    /// Optionally run a computation. 
    /// If the build fails return None otherwise retun Some<'a>.
    let optionMaybeM (ma:DocMonad<'a, 'userRes>) : DocMonad<'a option, 'userRes> = 
        combineM (fmapM Some ma)  (mreturn None)


    let optionalM (ma:DocMonad<'a, 'userRes>) : DocMonad<unit, 'userRes> = 
        combineM (fmapM ignore ma) (mreturn ())




    // ****************************************************
    // Logging

    let tellLine (msg:string) : DocMonad<unit, 'userRes> = 
        DocMonad <| fun sw _  _ -> 
            sw.WriteLine msg
            Ok ()

    // ****************************************************
    // Reader

    let ask () : DocMonad<DocBuildEnv, 'userRes> = 
        DocMonad <| fun _ _ env -> Ok env

    let asks (extract:DocBuildEnv -> 'a) : DocMonad<'a, 'userRes> = 
        DocMonad <| fun _ _ env -> Ok (extract env)


    let private assertDirectory (failMsg:string -> string) 
                                (path:string) : DocMonad<string, 'userRes> = 
        docMonad { 
            do! liftAssert (failMsg path) (Directory.Exists(path))
            return path
            }


    /// Note - this asserts that the Working directory path represents a 
    /// folder not a file.
    let askWorkingDirectory () : DocMonad<string, 'userRes> = 
        docMonad { 
            let! dir = asks (fun env -> env.WorkingDirectory)
            return! assertDirectory (sprintf "'Working' is not a directory: %s") dir
            }


    /// Note - this asserts that the Source directory path represents a 
    /// folder not a file.
    let askSourceDirectory () : DocMonad<string, 'userRes> = 
        docMonad { 
            let! dir = asks (fun env -> env.SourceDirectory) 
            return! assertDirectory (sprintf "'Source' is not a directory: %s") dir
        }




    /// Note - this asserts that the Include directory path represents a 
    /// folder not a file.
    let askIncludeDirectories () : DocMonad<string list, 'userRes> = 
        asks (fun env -> env.IncludeDirectories)



    /// Use with caution.
    /// Generally you might only want to update the 
    /// working directory
    let local (update:DocBuildEnv -> DocBuildEnv) 
              (ma:DocMonad<'a, 'userRes>) : DocMonad<'a, 'userRes> = 
        DocMonad <| fun sw res env -> 
            apply1 ma sw res (update env)


    let getResources () : DocMonad<AppResources<'userRes>, 'userRes> = 
        DocMonad <| fun _ res _ -> Ok res

    let getsResources (extract:AppResources<'userRes> -> 'a) : DocMonad<'a, 'userRes> = 
        DocMonad <| fun _ res _ -> Ok (extract res)

    let getUserResources () : DocMonad<'userRes, 'userRes> = 
        getsResources (fun res -> res.UserResources)

    let getsUserResources (extract:'userRes -> 'a) : DocMonad<'a, 'userRes> = 
        docMonad { 
            let! res = getUserResources ()
            return (extract res)
        }

            


 

    // ****************************************************
    // Monadic operations


    let whenM (cond:DocMonad<bool, 'userRes>) 
              (failMsg:string) 
              (successOp:unit -> DocMonad<'a, 'userRes>) = 
        docMonad { 
            let! ans = cond
            if ans then 
                let! res = successOp ()
                return res
            else docError failMsg |> ignore
            } 




    // liftM (which is fmap)
    let liftM (fn:'a -> 'x) (ma:DocMonad<'a, 'userRes>) : DocMonad<'x, 'userRes> = 
        fmapM fn ma

    let liftM2 (fn:'a -> 'b -> 'x) 
               (ma:DocMonad<'a, 'userRes>) 
               (mb:DocMonad<'b, 'userRes>) : DocMonad<'x, 'userRes> = 
        docMonad { 
            let! a = ma
            let! b = mb
            return (fn a b)
        }

    let liftM3 (fn:'a -> 'b -> 'c -> 'x) 
               (ma:DocMonad<'a, 'userRes>) 
               (mb:DocMonad<'b, 'userRes>) 
               (mc:DocMonad<'c, 'userRes>) : DocMonad<'x, 'userRes> = 
        docMonad { 
            let! a = ma
            let! b = mb
            let! c = mc
            return (fn a b c)
        }

    let liftM4 (fn:'a -> 'b -> 'c -> 'd -> 'x) 
               (ma:DocMonad<'a, 'userRes>) 
               (mb:DocMonad<'b, 'userRes>) 
               (mc:DocMonad<'c, 'userRes>) 
               (md:DocMonad<'d, 'userRes>) : DocMonad<'x, 'userRes> = 
        docMonad { 
            let! a = ma
            let! b = mb
            let! c = mc
            let! d = md
            return (fn a b c d)
        }


    let liftM5 (fn:'a -> 'b -> 'c -> 'd -> 'e -> 'x) 
               (ma:DocMonad<'a, 'userRes>) 
               (mb:DocMonad<'b, 'userRes>) 
               (mc:DocMonad<'c, 'userRes>) 
               (md:DocMonad<'d, 'userRes>) 
               (me:DocMonad<'e, 'userRes>) : DocMonad<'x, 'userRes> = 
        docMonad { 
            let! a = ma
            let! b = mb
            let! c = mc
            let! d = md
            let! e = me
            return (fn a b c d e)
        }

    let liftM6 (fn:'a -> 'b -> 'c -> 'd -> 'e -> 'f -> 'x) 
               (ma:DocMonad<'a, 'userRes>) 
               (mb:DocMonad<'b, 'userRes>) 
               (mc:DocMonad<'c, 'userRes>) 
               (md:DocMonad<'d, 'userRes>) 
               (me:DocMonad<'e, 'userRes>) 
               (mf:DocMonad<'f, 'userRes>) : DocMonad<'x, 'userRes> = 
        docMonad { 
            let! a = ma
            let! b = mb
            let! c = mc
            let! d = md
            let! e = me
            let! f = mf
            return (fn a b c d e f)
        }


    let tupleM2 (ma:DocMonad<'a, 'userRes>) 
                (mb:DocMonad<'b, 'userRes>) : DocMonad<'a * 'b, 'userRes> = 
        liftM2 (fun a b -> (a,b)) ma mb

    let tupleM3 (ma:DocMonad<'a, 'userRes>) 
                (mb:DocMonad<'b, 'userRes>) 
                (mc:DocMonad<'c, 'userRes>) : DocMonad<'a * 'b * 'c, 'userRes> = 
        liftM3 (fun a b c -> (a,b,c)) ma mb mc

    let tupleM4 (ma:DocMonad<'a, 'userRes>) 
                (mb:DocMonad<'b, 'userRes>) 
                (mc:DocMonad<'c, 'userRes>) 
                (md:DocMonad<'d, 'userRes>) : DocMonad<'a * 'b * 'c * 'd, 'userRes> = 
        liftM4 (fun a b c d -> (a,b,c,d)) ma mb mc md

    let tupleM5 (ma:DocMonad<'a, 'userRes>) 
                (mb:DocMonad<'b, 'userRes>) 
                (mc:DocMonad<'c, 'userRes>) 
                (md:DocMonad<'d, 'userRes>) 
                (me:DocMonad<'e, 'userRes>) : DocMonad<'a * 'b * 'c * 'd * 'e, 'userRes> = 
        liftM5 (fun a b c d e -> (a,b,c,d,e)) ma mb mc md me

    let tupleM6 (ma:DocMonad<'a, 'userRes>) 
                (mb:DocMonad<'b, 'userRes>) 
                (mc:DocMonad<'c, 'userRes>) 
                (md:DocMonad<'d, 'userRes>) 
                (me:DocMonad<'e, 'userRes>) 
                (mf:DocMonad<'f, 'userRes>) : DocMonad<'a * 'b * 'c * 'd * 'e * 'f, 'userRes> = 
        liftM6 (fun a b c d e f -> (a,b,c,d,e,f)) ma mb mc md me mf

    let pipeM2 (ma:DocMonad<'a, 'userRes>) 
               (mb:DocMonad<'b, 'userRes>) 
               (fn:'a -> 'b -> 'x) : DocMonad<'x, 'userRes> = 
        liftM2 fn ma mb

    let pipeM3 (ma:DocMonad<'a, 'userRes>) 
               (mb:DocMonad<'b, 'userRes>) 
               (mc:DocMonad<'c, 'userRes>) 
               (fn:'a -> 'b -> 'c -> 'x) : DocMonad<'x, 'userRes> = 
        liftM3 fn ma mb mc

    let pipeM4 (ma:DocMonad<'a, 'userRes>) 
               (mb:DocMonad<'b, 'userRes>) 
               (mc:DocMonad<'c, 'userRes>) 
               (md:DocMonad<'d, 'userRes>) 
               (fn:'a -> 'b -> 'c -> 'd -> 'x) : DocMonad<'x, 'userRes> = 
        liftM4 fn ma mb mc md

    let pipeM5 (ma:DocMonad<'a, 'userRes>) 
               (mb:DocMonad<'b, 'userRes>) 
               (mc:DocMonad<'c, 'userRes>) 
               (md:DocMonad<'d, 'userRes>) 
               (me:DocMonad<'e, 'userRes>) 
               (fn:'a -> 'b -> 'c -> 'd -> 'e ->'x) : DocMonad<'x, 'userRes> = 
        liftM5 fn ma mb mc md me

    let pipeM6 (ma:DocMonad<'a, 'userRes>) 
               (mb:DocMonad<'b, 'userRes>) 
               (mc:DocMonad<'c, 'userRes>) 
               (md:DocMonad<'d, 'userRes>) 
               (me:DocMonad<'e, 'userRes>) 
               (mf:DocMonad<'f, 'userRes>) 
               (fn:'a -> 'b -> 'c -> 'd -> 'e -> 'f -> 'x) : DocMonad<'x, 'userRes> = 
        liftM6 fn ma mb mc md me mf

    /// Left biased choice, if ``ma`` succeeds return its result, otherwise try ``mb``.
    let altM (ma:DocMonad<'a, 'userRes>) (mb:DocMonad<'a, 'userRes>) : DocMonad<'a, 'userRes> = 
        combineM ma mb


    /// Haskell Applicative's (<*>)
    let apM (mf:DocMonad<'a ->'b, 'userRes>) (ma:DocMonad<'a, 'userRes>) : DocMonad<'b, 'userRes> = 
        docMonad { 
            let! fn = mf
            let! a = ma
            return (fn a) 
        }



    /// Perform two actions in sequence. 
    /// Ignore the results of the second action if both succeed.
    let seqL (ma:DocMonad<'a, 'userRes>) (mb:DocMonad<'b, 'userRes>) : DocMonad<'a, 'userRes> = 
        docMonad { 
            let! a = ma
            let! b = mb
            return a
        }

    /// Perform two actions in sequence. 
    /// Ignore the results of the first action if both succeed.
    let seqR (ma:DocMonad<'a, 'userRes>) (mb:DocMonad<'b, 'userRes>) : DocMonad<'b, 'userRes> = 
        docMonad { 
            let! a = ma
            let! b = mb
            return b
        }





    let optionToFailM (errMsg:string)
                      (ma:DocMonad<'a option, 'userRes>) : DocMonad<'a, 'userRes> = 
        bindM ma (fun opt -> 
                    match opt with
                    | Some ans -> mreturn ans
                    | None -> docError errMsg)


    let kleisliL (mf:'a -> DocMonad<'b, 'userRes>)
                 (mg:'b -> DocMonad<'c, 'userRes>)
                 (source:'a) : DocMonad<'c, 'userRes> = 
        docMonad { 
            let! b = mf source
            let! c = mg b
            return c
        }

    /// Flipped kleisliL
    let kleisliR (mf:'b -> DocMonad<'c, 'userRes>)
                 (mg:'a -> DocMonad<'b, 'userRes>)
                 (source:'a) : DocMonad<'c, 'userRes> = 
        docMonad { 
            let! b = mg source
            let! c = mf b
            return c
        }


    // ****************************************************
    // Execute 'builtin' processes 
    // (Respective applications must be installed)

    type private ProcessOptions = 
        { WorkingDirectory: string 
          ExecutableName: string 
        }

    let private getProcessOptions 
                    (findExe:AppResources<'userRes> -> string) : DocMonad<ProcessOptions, 'userRes> = 
        docMonad { 
            let! exe = getsResources findExe
            let! cwd = askWorkingDirectory ()
            return { WorkingDirectory = cwd
                   ; ExecutableName = exe}
        }
    
    let private shellExecute (findExe:AppResources<'userRes> -> string)
                             (args:CmdOpt list) : DocMonad<string, 'userRes> = 
        docMonad { 
            let! options = getProcessOptions findExe
            let! ans = 
                liftOperationResult "shellExecute" 
                    <| fun _ -> SimpleInvoke.runProcess options.WorkingDirectory 
                                                        options.ExecutableName
                                                            args
            return ans.StdOut
            }
        
    let execGhostscript (args:CmdOpt list) : DocMonad<string, 'userRes> = 
        shellExecute (fun res -> res.GhostscriptExe) args

    let execPandoc (args:CmdOpt list) : DocMonad<string, 'userRes> = 
        shellExecute (fun res -> res.PandocExe) args

    let execPdftk (args:CmdOpt list) : DocMonad<string, 'userRes> = 
        shellExecute (fun res -> res.PdftkExe) args

    // ****************************************************
    // Recursive functions


    /// Implemented in CPS 
    let mapM (mf: 'a -> DocMonad<'b, 'userRes>) 
             (source:'a list) : DocMonad<'b list, 'userRes> = 
        DocMonad <| fun sw res env -> 
            let rec work ac ys fk sk = 
                match ys with
                | [] -> sk (List.rev ac)
                | z :: zs -> 
                    match apply1 (mf z) sw res env with
                    | Error msg -> fk msg
                    | Ok ans -> 
                        work ac zs fk (fun acs ->
                        sk (ans::acs))
            work [] source (fun msg -> Error msg) (fun ans -> Ok ans)

    /// Flipped mapM
    let forM (source:'a list) 
             (mf: 'a -> DocMonad<'b, 'userRes>) : DocMonad<'b list, 'userRes> = 
        mapM mf source

    /// Forgetful mapM
    let mapMz (mf: 'a -> DocMonad<'b, 'userRes>) 
              (source:'a list) : DocMonad<unit, 'userRes> = 
        DocMonad <| fun sw res env -> 
            let rec work ys cont = 
                match ys with
                | [] -> cont (Ok ())
                | z :: zs -> 
                    match apply1 (mf z) sw res env with
                    | Error msg -> cont (Error msg)
                    | Ok ans -> work zs cont
            work source id

    /// Flipped mapMz
    let forMz (source:'a list) 
              (mf: 'a -> DocMonad<'b, 'userRes>) : DocMonad<unit, 'userRes> = 
        mapMz mf source


    /// Implemented in CPS 
    let mapiM (mf:int -> 'a -> DocMonad<'b, 'userRes>) 
              (source:'a list) : DocMonad<'b list, 'userRes> = 
        DocMonad <| fun sw res env -> 
            let rec work ac n ys fk sk = 
                match ys with
                | [] -> sk (List.rev ac)
                | z :: zs -> 
                    match apply1 (mf n z) sw res env with
                    | Error msg -> fk msg
                    | Ok ans -> 
                        work ac (n+1) zs fk (fun acs ->
                        sk (ans::acs))
            work [] 0 source (fun msg -> Error msg) (fun ans -> Ok ans)

    /// Flipped mapMi
    let foriM (source:'a list) 
              (mf: int -> 'a -> DocMonad<'b, 'userRes>)  : DocMonad<'b list, 'userRes> = 
        mapiM mf source

    /// Forgetful mapiM
    let mapiMz (mf: int -> 'a -> DocMonad<'b, 'userRes>) 
              (source:'a list) : DocMonad<unit, 'userRes> = 
        DocMonad <| fun sw res env -> 
            let rec work n ys cont = 
                match ys with
                | [] -> cont (Ok ())
                | z :: zs -> 
                    match apply1 (mf n z) sw res env with
                    | Error msg -> cont (Error msg)
                    | Ok ans -> work (n+1) zs cont
            work 0 source id

    /// Flipped mapiMz
    let foriMz (source:'a list) 
               (mf: int -> 'a -> DocMonad<'b, 'userRes>) : DocMonad<unit, 'userRes> = 
        mapiMz mf source


    /// Implemented in CPS 
    let firstOfM (actions: DocMonad<'a, 'userRes> list) : DocMonad<'a, 'userRes> = 
        DocMonad <| fun sw res env -> 
            let rec work ops fk sk = 
                match ops with
                | [] -> fk "firstOfM - no successes"
                | ma :: rest -> 
                    match apply1 ma sw res env with
                    | Ok ans -> sk ans
                    | Error _ -> 
                        work rest fk sk
            work actions (fun msg -> Error msg) (fun ans -> Ok ans)

    // ****************************************************
    // Operators

    // ****************************************************
    // Errors

    /// Operator for swapError
    let ( <?&> ) (msg:string) (ma:DocMonad<'a, 'userRes>) : DocMonad<'a, 'userRes> = 
        swapError msg ma

    /// Operator for flip swapError
    let ( <&?> ) (ma:DocMonad<'a, 'userRes>) (msg:string) : DocMonad<'a, 'userRes> = 
        swapError msg ma


    // ****************************************************
    // Monadic operations

    /// Bind operator
    let ( >>= ) (ma:DocMonad<'a, 'userRes>) 
              (fn:'a -> DocMonad<'b, 'userRes>) : DocMonad<'b, 'userRes> = 
        docMonad.Bind(ma,fn)

    /// Flipped Bind operator
    let ( =<< ) (fn:'a -> DocMonad<'b, 'userRes>) 
              (ma:DocMonad<'a, 'userRes>) : DocMonad<'b, 'userRes> = 
        docMonad.Bind(ma,fn)


    /// Operator for fmap.
    let ( |>> ) (ma:DocMonad<'a, 'userRes>) (fn:'a -> 'b) : DocMonad<'b, 'userRes> = 
        fmapM fn ma

    /// Flipped fmap.
    let ( <<| ) (fn:'a -> 'b) (ma:DocMonad<'a, 'userRes>) : DocMonad<'b, 'userRes> = 
        fmapM fn ma

    /// Operator for altM
    let ( <|> ) (ma:DocMonad<'a, 'userRes>) 
               (mb:DocMonad<'a, 'userRes>) : DocMonad<'a, 'userRes> = 
        altM ma mb 


    /// Operator for apM
    let ( <**> ) (ma:DocMonad<'a -> 'b, 'userRes>) 
               (mb:DocMonad<'a, 'userRes>) : DocMonad<'b, 'userRes> = 
        apM ma mb

    /// Operator for fmapM
    let ( <&&> ) (fn:'a -> 'b) (ma:DocMonad<'a, 'userRes>) : DocMonad<'b, 'userRes> = 
        fmapM fn ma



    /// Operator for seqL
    let (.>>) (ma:DocMonad<'a, 'userRes>) 
              (mb:DocMonad<'b, 'userRes>) : DocMonad<'a, 'userRes> = 
        seqL ma mb

    /// Operator for seqR
    let (>>.) (ma:DocMonad<'a, 'userRes>) 
              (mb:DocMonad<'b, 'userRes>) : DocMonad<'b, 'userRes> = 
        seqR ma mb



    /// Operator for kleisliL
    let ( >=> ) (mf : 'a -> DocMonad<'b, 'userRes>)
              (mg : 'b -> DocMonad<'c, 'userRes>)
              (source:'a) : DocMonad<'c, 'userRes> = 
        kleisliL mf mg source


    /// Operator for kleisliR
    let ( <=< ) (mf : 'b -> DocMonad<'c, 'userRes>)
              (mg : 'a -> DocMonad<'b, 'userRes>)
              (source:'a) : DocMonad<'c, 'userRes> = 
        kleisliR mf mg source



    let ignoreM (ma:DocMonad<'a, 'userRes>) : DocMonad<unit, 'userRes> = 
        ma |>> ignore

