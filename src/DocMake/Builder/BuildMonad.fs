// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

module DocMake.Builder.BuildMonad

open System.Text

open DocMake.Base.Common



type Env = 
    { WorkingDirectory: string
      PrintQuality: PrintQuality
      PdfQuality: PdfPrintQuality }



/// State is a histogram as the state that holds a counter for each 
/// file type. e.g a map of type: Map<string,int>


/// Map: extension -> count
type FileTypeHistogram = Map<string,int>

let fileTypeHistogramNext (extension:string) (histo:FileTypeHistogram) : FileTypeHistogram * int = 
    match Map.tryFind extension histo with
    | None -> (Map.add extension 1 histo, 1)
    | Some i -> (Map.add extension (i+1) histo, i+ 1)

type State = FileTypeHistogram


type FailMsg = string

type BuildError = BuildError of string * list<BuildError>

let getErrorLog (err:BuildError) : string = 
    let writeLine (depth:int) (str:string) (sb:StringBuilder) : StringBuilder = 
        let line = sprintf "%s %s" (String.replicate depth "*") str
        sb.AppendLine(line)
    let rec work (e1:BuildError) (depth:int) (sb:StringBuilder) : StringBuilder  = 
        match e1 with
        | BuildError (s,[]) -> writeLine depth s sb
        | BuildError (s,xs) ->
            let sb1 = writeLine depth s sb
            List.fold (fun buf branch -> work branch (depth+1) buf) sb1 xs
    work err 0 (new StringBuilder()) |> fun sb -> sb.ToString()

/// Create a BuildError
let buildError (errMsg:string) : BuildError = 
    BuildError(errMsg, [])

let concatBuildErrors (errMsg:string) (failures:BuildError list) : BuildError = 
    BuildError(errMsg,failures)
    
type BuildMonadResult<'a> =
    | BmErr of BuildError
    | BmOk of State * 'a




// Note - keeping log in a StringBuilder means we only see it "at the end",
// any direct console writes are visible before the log is shown.
type BuildMonad<'res,'a> = 
    private BuildMonad of ((Env * 'res)  -> State -> BuildMonadResult<'a>)

let inline private apply1 (ma : BuildMonad<'res,'a>) (handle:Env * 'res) (st:State) : BuildMonadResult<'a> = 
    let (BuildMonad f) = ma in f handle st

// Return in the BuildMonad
let inline breturn (x:'a) : BuildMonad<'res,'a> = 
    BuildMonad (fun _ st -> BmOk (st,x))

let private failM : BuildMonad<'res,'a> = 
    BuildMonad (fun _ st -> BmErr (buildError "failM"))


let inline private bindM (ma:BuildMonad<'res,'a>) (f : 'a -> BuildMonad<'res,'b>) : BuildMonad<'res,'b> =
    BuildMonad <| fun res st0 -> 
        match apply1 ma res st0 with
        | BmErr msg -> BmErr msg
        | BmOk (st1,a) -> apply1 (f a) res st1


let inline private altM (ma:BuildMonad<'res,'a>) (mb:BuildMonad<'res,'a>) : BuildMonad<'res,'a> =
    BuildMonad <| fun res st0 -> 
        match apply1 ma res st0 with 
        | BmErr stk1 -> 
            match apply1 mb res st0 with
            | BmErr stk2 -> BmErr (concatBuildErrors "altM" [stk1;stk2])
            | BmOk (st2,b) -> BmOk (st2,b)
        | BmOk (st1,a) -> BmOk (st1, a)

/// This is Haskell's (>>).
let inline private combineM (ma:BuildMonad<'res,unit>) (mb:BuildMonad<'res,'b>) : BuildMonad<'res,'b> = 
    BuildMonad <| fun res st0 -> 
        match apply1 ma res st0 with
        | BmErr msg -> BmErr msg
        | BmOk (st1,_) -> 
            match apply1 mb res st1 with
            | BmErr msg -> BmErr msg
            | BmOk (st2,b) -> BmOk (st2, b)

let inline private delayM (fn:unit -> BuildMonad<'res,'a>) : BuildMonad<'res,'a> = 
    bindM (breturn ()) fn 


type BuildMonadBuilder() = 
    member self.Return x        = breturn x
    member self.Bind (p,f)      = bindM p f
    member self.Zero ()         = failM
    member self.Delay fn        = delayM fn
    member self.Combine (p,q)   = combineM p q

let (buildMonad:BuildMonadBuilder) = new BuildMonadBuilder()


// Common monadic operations
let fmapM (fn:'a -> 'b) (ma:BuildMonad<'res,'a>) : BuildMonad<'res,'b> = 
    BuildMonad <| fun res st0 ->
        match apply1 ma res st0  with
        | BmErr msg -> BmErr msg
        | BmOk (st1,a) -> BmOk (st1, fn a)


let mapM (fn:'a -> BuildMonad<'res,'b>) (xs:'a list) : BuildMonad<'res,'b list> =
    BuildMonad <| fun res state -> 
        let rec work ac (st0:State) ys = 
            match ys with
            | [] -> BmOk (st0, List.rev ac)
            | z :: zs ->
                match apply1 (fn z) res st0 with
                | BmErr msg -> BmErr msg 
                | BmOk(st1, a) -> work (a::ac) st1 zs
        work [] state xs

let forM (xs:'a list) (fn:'a -> BuildMonad<'res,'b>) : BuildMonad<'res,'b list> = 
    mapM fn xs


let mapMz (fn:'a -> BuildMonad<'res,'b>) (xs:'a list) : BuildMonad<'res,unit> =
    BuildMonad <| fun res state -> 
        let rec work st0 ys = 
            match ys with
            | [] -> BmOk (st0, ())
            | z :: zs ->
                match apply1 (fn z) res st0 with
                | BmErr msg -> BmErr msg
                | BmOk (st1,_) -> work st1 zs
        work state xs


let forMz (xs:'a list) (fn:'a -> BuildMonad<'res,'b>) : BuildMonad<'res,unit> = 
    mapMz fn xs



//let traverseM (fn: 'a -> BuildMonad<'res,'b>) (source:seq<'a>) : BuildMonad<'res,seq<'b>> = 
//    BuildMonad <| fun res sbuf state ->
//        Seq.mapFold (fun ac x -> 
//                        match ac with
//                        
//                        let (st1,ans  = apply1 (fn x) res sbuf in ac) 
//                     (GoodTrav state ())
//                     source 


//let traverseMz (fn: 'a -> BuildMonad<'res,'b>) (source:seq<'a>) : BuildMonad<'res,unit> = 
//    BuildMonad <| fun res sbuf ->
//        try 
//            Seq.fold (fun ac x -> 
//                        let ans  = apply1Ex (fn x) res sbuf in ac) 
//                     (BmOk ())
//                     source 
//        with
//        | ex -> Err (ex.Message)



let mapiM (fn:int -> 'a -> BuildMonad<'res,'b>) (xs:'a list) : BuildMonad<'res,'b list> = 
    BuildMonad <| fun res state -> 
        let rec work ac ix st0 ys = 
            match ys with
            | [] -> BmOk (st0,List.rev ac)
            | z :: zs ->
                match apply1 (fn ix z) res st0 with
                | BmErr msg -> BmErr msg
                | BmOk (st1,a) -> work (a::ac) (ix+1) st1 zs
        work [] 0 state xs

let mapiMz (fn:int -> 'a -> BuildMonad<'res,'b>) (xs:'a list) : BuildMonad<'res,unit> =
    BuildMonad <| fun res state -> 
        let rec work ix st0 ys = 
            match ys with
            | [] -> BmOk (st0, ())
            | z :: zs ->
                match apply1 (fn ix z) res st0 with
                | BmErr msg -> BmErr msg
                | BmOk (st1, _) -> work (ix+1) st1 zs
        work 0 state xs

/// Flipped mapiM
let foriM (xs:'a list) (fn:int -> 'a -> BuildMonad<'res,'b>) : BuildMonad<'res,'b list> =
    mapiM fn xs

/// Flipped mapiMz
let foriMz (xs:'a list) (fn:int -> 'a -> BuildMonad<'res,'b>) : BuildMonad<'res,unit> =
    mapiMz fn xs

// Alternative combinators would be useful...

// *********************************************************
// *********************************************************


// Symbolic aliases / combinators
// Delegate to FParsec for naming where we can.

// fmapM - aka Haskell's (<$>)
let (<<|) (fn:'a -> 'b)  (ma: BuildMonad<'res,'a>) : BuildMonad<'res,'b> = 
    fmapM fn ma


// Flipped fmapM 
let (|>>) (ma: BuildMonad<'res,'a>) (fn:'a -> 'b) : BuildMonad<'res,'b> = 
    fmapM fn ma


// Aka Haskell's (*>)
let (>>.) (ma: BuildMonad<'res,'a>) (mb: BuildMonad<'res,'b>) : BuildMonad<'res,'b> = 
    buildMonad { 
        let! _ = ma
        let! b = mb
        return b 
    }

// Aka Haskell's (<*)
let (.>>) (ma: BuildMonad<'res,'a>) (mb: BuildMonad<'res,'b>) : BuildMonad<'res,'a> =  
    buildMonad { 
        let! a = ma
        let! _ = mb
        return a 
    }


// Alt
let (<|>) (ma:BuildMonad<'res,'a>) (mb:BuildMonad<'res,'a>) : BuildMonad<'res,'a> = 
    altM ma mb

/// Monadic bind
let (>>=) (ma: BuildMonad<'res,'a>) (k: 'a -> BuildMonad<'res,'b>) : BuildMonad<'res,'b> = 
    bindM ma k

/// Flipped monadic bind
let (=<<) (k: 'a -> BuildMonad<'res,'b>) (ma: BuildMonad<'res,'a>) : BuildMonad<'res,'b> = 
    bindM ma k


// *********************************************************
// *********************************************************

// BuildMonad operations
let runBuildMonad (env:Env) (handle:'res) (stateZero:State) (ma:BuildMonad<'res,'a>) : BuildMonadResult<'a>= 
    match ma with 
    | BuildMonad fn ->  fn (env,handle) stateZero


// TODO need a simple way to run things
// In Haskell the `eval` prefix is closes to "run a cmputation, return (just) the answer"

let evalBuildMonad (env:Env) (handle:'res) (finalizer:'res -> unit) (stateZero:State)  (ma:BuildMonad<'res,'a>) : 'a = 
    let ans = runBuildMonad env handle stateZero ma
    finalizer handle
    match ans with
    | BmErr stk -> failwith (getErrorLog stk)
    | BmOk (_,a) -> a



let consoleRun (env:Env) (handle:'res) (finalizer:'res -> unit) (ma:BuildMonad<'res,'a>) : 'a = 
    let stateZero : State = Map.empty
    evalBuildMonad env handle finalizer stateZero ma



let withUserHandle (handle:'uhandle) (finalizer:'uhandle -> unit) (ma:BuildMonad<'uhandle,'a>) : BuildMonad<'res,'a> = 
    BuildMonad <| fun (env,_) st0 -> 
        let ans = runBuildMonad env handle st0 ma
        finalizer handle
        ans

let throwError (msg:string) : BuildMonad<'res,'a> = 
    BuildMonad <| fun _ _ -> BmErr (buildError msg)


let swapError (msg:string) (ma:BuildMonad<'res,'a>) : BuildMonad<'res,'a> = 
    BuildMonad <| fun (env,res) st0 -> 
        match apply1 ma (env, res) st0 with
        | BmErr (BuildError (_,stk)) -> BmErr (BuildError (msg,stk))
        | BmOk (pos1,a) -> BmOk (pos1,a)


let (<&?>) (ma:BuildMonad<'res,'a>) (msg:string) : BuildMonad<'res,'a> = 
    swapError msg ma

let (<?&>) (msg:string) (ma:BuildMonad<'res,'a>) : BuildMonad<'res,'a> = 
    swapError msg ma


/// Execute an action that may throw an exception, capture the exception 
/// as a failure in the monad.
let attempt (ma: BuildMonad<'res,'a>) : BuildMonad<'res,'a> = 
    BuildMonad <| fun (env,res) st0 -> 
        try
            apply1 ma (env, res) st0
        with
        | _ -> BmErr (buildError "attempt failed")



/// Execute an arbitrary FSharp action that may use IO, throw an exception, etc.
/// Capture any failure within the BuildMonad.
/// (It seems like this proc needs to be guarded with a thunk)
let executeIO (operation:unit -> 'a) : BuildMonad<'res,'a> = 
    BuildMonad <| fun _ st0 -> 
    try 
        let ans = operation () in BmOk (st0, ans)
    with
    | ex -> BmErr (buildError (sprintf "executeIO: %s" ex.Message))



/// Note unit param to avoid value restriction.
let askU () : BuildMonad<'res,'res> = 
    BuildMonad <| fun (_,res) st0 -> BmOk (st0, res)

let asksU (project:'res -> 'a) : BuildMonad<'res,'a> = 
    BuildMonad <| fun (_,res) st0 -> BmOk (st0, project res)

let localU (modify:'res -> 'res) (ma:BuildMonad<'res,'a>) : BuildMonad<'res,'a> = 
    BuildMonad <| fun (env,res) st0 -> apply1 ma (env, modify res) st0

    // PrintQuality
let askEnv () : BuildMonad<'res,Env> = 
    BuildMonad <| fun (env,_) st0 -> BmOk (st0, env)

let asksEnv (project:Env -> 'a) : BuildMonad<'res,'a> = 
    BuildMonad <| fun (env,_) st0 -> BmOk (st0, project env)

let localEnv (modify:Env -> Env) (ma:BuildMonad<'res,'a>) : BuildMonad<'res,'a> = 
    BuildMonad <| fun (env,res) st0 -> apply1 ma (modify env, res) st0

/// The MakeName function is in State rather than Env, but we only provide an API
/// to run it within a context (cf. the Reader monad's local), rather than reset it 
/// imperatively.



// TODO - if this took a file extension it might simplify things
let freshFileName (extension:string) : BuildMonad<'res, string> = 
    BuildMonad <| fun (env,_) st0 -> 
        let (st1,i) = fileTypeHistogramNext extension st0
        let name1 = sprintf "temp%03i.%s" i extension
        let outPath = System.IO.Path.Combine(env.WorkingDirectory,name1)
        BmOk (st1, outPath)
