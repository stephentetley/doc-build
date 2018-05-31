module DocMake.Builder.BuildMonad

open System.Text

open DocMake.Base.Common


type FailMsg = string

type Answer<'a> =
    | Err of FailMsg
    | Ok of 'a

// There's a design issue that we probably won't explore, but it is very 
// interesting (although complicated).
// With continuations might enable us to have restartable builds, this could 
// get us back to an idea of workflows that we have ignored for simplicity 
// reasons.


/// TODO - adding "working directory" and a counter would be useful
/// We could readily generate temp files

type Env = 
    { WorkingDirectory: string
      PrintQuality: DocMakePrintQuality
      PdfQuality: PdfPrintSetting }



type State = 
    { MakeName: int -> string
      NameIndex: int }

let private incrNameIndex (st:State) : State = 
    let i = st.NameIndex in {st with NameIndex = i + 1 }



// Note - keeping log in a StringBuilder means we only see it "at the end",
// any direct console writes are visible before the log is shown.
type BuildMonad<'res,'a> = 
    private BuildMonad of ((Env * 'res)  -> State -> (State * Answer<'a>))

let inline private apply1 (ma : BuildMonad<'res,'a>) (handle:Env * 'res) (st:State) : State *  Answer<'a> = 
    let (BuildMonad f) = ma in f handle st

// Return in the BuildMonad
let inline breturn (x:'a) : BuildMonad<'res,'a> = 
    BuildMonad (fun _ st -> st, Ok x)

let private failM : BuildMonad<'res,'a> = 
    BuildMonad (fun _ st -> st, Err "failM")


let inline private bindM (ma:BuildMonad<'res,'a>) (f : 'a -> BuildMonad<'res,'b>) : BuildMonad<'res,'b> =
    BuildMonad <| fun res st0 -> 
        let st1, ans = apply1 ma res st0
        match ans with
        | Err s -> (st1, Err s)
        | Ok a -> apply1 (f a) res st1


let inline private altM (ma:BuildMonad<'res,'a>) (mb:BuildMonad<'res,'a>) : BuildMonad<'res,'a> =
    BuildMonad <| fun res st0 -> 
        let st1, ans = apply1 ma res st0
        match ans with 
        | Err s -> apply1 ma res st1
        | Ok a -> st1, Ok a

type BuildMonadBuilder() = 
    member self.Return x = breturn x
    member self.Bind (p,f) = bindM p f
    member self.Zero () = failM

let (buildMonad:BuildMonadBuilder) = new BuildMonadBuilder()


// Common monadic operations
let fmapM (fn:'a -> 'b) (ma:BuildMonad<'res,'a>) : BuildMonad<'res,'b> = 
    BuildMonad <| fun res st0 ->
        let (st1,ans) =  apply1 ma res st0 
        match ans with
        | Err msg -> (st1, Err msg)
        | Ok a -> (st1, Ok <| fn a)


let mapM (fn:'a -> BuildMonad<'res,'b>) (xs:'a list) : BuildMonad<'res,'b list> =
    BuildMonad <| fun res state -> 
        let rec work ac (st0:State) ys = 
            match ys with
            | [] -> (st0, Ok <| List.rev ac)
            | z :: zs ->
                let (st1,ans) = apply1 (fn z) res st0
                match ans with
                | Err msg -> st1, Err msg 
                | Ok a -> work (a::ac) st1 zs
        work [] state xs

let forM (xs:'a list) (fn:'a -> BuildMonad<'res,'b>) : BuildMonad<'res,'b list> = 
    mapM fn xs


let mapMz (fn:'a -> BuildMonad<'res,'b>) (xs:'a list) : BuildMonad<'res,unit> =
    BuildMonad <| fun res state -> 
        let rec work st0 ys = 
            match ys with
            | [] -> (st0, Ok ())
            | z :: zs ->
                let (st1,ans) = apply1 (fn z) res st0
                match ans with
                | Err msg -> (st1, Err msg)
                | Ok _ -> work st1 zs
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
//                     (Ok ())
//                     source 
//        with
//        | ex -> Err (ex.Message)



let mapiM (fn:int -> 'a -> BuildMonad<'res,'b>) (xs:'a list) : BuildMonad<'res,'b list> = 
    BuildMonad <| fun res state -> 
        let rec work ac ix st0 ys = 
            match ys with
            | [] -> (st0, Ok <| List.rev ac)
            | z :: zs ->
                let (st1,ans) = apply1 (fn ix z) res st0
                match ans with
                | Err msg -> (st1, Err msg)
                | Ok a -> work (a::ac) (ix+1) st1 zs
        work [] 0 state xs

let mapiMz (fn:int -> 'a -> BuildMonad<'res,'b>) (xs:'a list) : BuildMonad<'res,unit> =
    BuildMonad <| fun res state -> 
        let rec work ix st0 ys = 
            match ys with
            | [] -> (st0, Ok ())
            | z :: zs ->
                let (st1,ans) =  apply1 (fn ix z) res st0
                match ans with
                | Err msg -> (st1, Err msg)
                | Ok _ -> work (ix+1) st1 zs
        work 0 state xs

let foriM (fn:int -> 'a -> BuildMonad<'res,'b>) (xs:'a list) : BuildMonad<'res,'b list> =
    mapiM fn xs

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


// *********************************************************
// *********************************************************

// BuildMonad operations
let runBuildMonad (env:Env) (handle:'res) (stateZero:State) (ma:BuildMonad<'res,'a>) : State * Answer<'a>= 
    match ma with 
    | BuildMonad fn -> let (s1,ans) = fn (env,handle) stateZero in (s1,ans)


// TODO need a simple way to run things
// In Haskell the `eval` prefix is closes to "run a cmputation, return (just) the answer"

let evalBuildMonad (env:Env) (handle:'res) (finalizer:'res -> unit) (stateZero:State)  (ma:BuildMonad<'res,'a>) : 'a = 
    let _, ans = runBuildMonad env handle stateZero ma
    finalizer handle
    match ans with
    | Err msg -> failwith msg
    | Ok a -> a

let consoleRun (env:Env) (ma:BuildMonad<unit,'a>) : 'a = 
    let stateZero : State = 
        { MakeName = sprintf "temp%03i" 
          NameIndex = 1 }
    evalBuildMonad env () (fun _ -> ()) stateZero ma



let withUserHandle (handle:'uhandle) (finalizer:'uhandle -> unit) (ma:BuildMonad<'uhandle,'a>) : BuildMonad<'res,'a> = 
    BuildMonad <| fun (env,_) st0 -> 
        let (st1,ans) = runBuildMonad env handle st0 ma
        finalizer handle
        (st1,ans)

let throwError (msg:string) : BuildMonad<'res,'a> = 
    BuildMonad <| fun _ st0 -> 
        (st0, Err msg)




/// Execute an FSharp action that may use IO, throw an exception...
/// Capture any failure within the BuildMonad.
/// (It seems like this proc needs to be guarded with a thunk)
let executeIO (operation:unit -> 'a) : BuildMonad<'res,'a> = 
    BuildMonad <| fun _ st0 -> 
    try 
        let ans = operation () in (st0, Ok ans)
    with
    | ex -> (st0, Err ex.Message)



/// Note unit param to avoid value restriction.
let askU () : BuildMonad<'res,'res> = 
    BuildMonad <| fun (_,res) st0 -> (st0, Ok res)

let asksU (project:'res -> 'a) : BuildMonad<'res,'a> = 
    BuildMonad <| fun (_,res) st0 -> (st0, Ok (project res))

let localU (modify:'res -> 'res) (ma:BuildMonad<'res,'a>) : BuildMonad<'res,'a> = 
    BuildMonad <| fun (env,res) st0 -> apply1 ma (env, modify res) st0

    // PrintQuality
let askEnv () : BuildMonad<'res,Env> = 
    BuildMonad <| fun (env,_) st0 -> (st0, Ok env)

let asksEnv (project:Env -> 'a) : BuildMonad<'res,'a> = 
    BuildMonad <| fun (env,_) st0 -> (st0, Ok (project env))

let localEnv (modify:Env -> Env) (ma:BuildMonad<'res,'a>) : BuildMonad<'res,'a> = 
    BuildMonad <| fun (env,res) st0 -> apply1 ma (modify env, res) st0

/// The MakeName function is in State rather than Env, but we only provide an API
/// to run it within a context (cf. the Reader monad's local), rather than reset it 
/// imperatively.


/// Note the file number increases with each file generated, not each type of file generated.
let withNameGen (namer:int -> string) (ma:BuildMonad<'res,'a>): BuildMonad<'res,'a> =
    BuildMonad <| fun (env,res) st0 -> 
        let fun1 = st0.MakeName
        let (s1,ans) = apply1 ma (env,res) {st0 with MakeName = namer}
        ({s1 with MakeName = fun1}, ans)


let freshFileName () : BuildMonad<'res, string> = 
    BuildMonad <| fun (env,_) st0 -> 
        let i = st0.NameIndex
        let name1 = st0.MakeName i
        let outPath = System.IO.Path.Combine(env.WorkingDirectory,name1)
        (incrNameIndex st0, Ok outPath)
