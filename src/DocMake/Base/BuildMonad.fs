module DocMake.Base.BuildMonad

open System.Text

open DocMake.Base.Common


/// Document has a Phantom Type so we can distinguish between different types 
/// (Word, Excel, Pdf, ...)
/// Maybe we ought to store whether a file has been derived in the build process
/// (and so deleteable)... 
type Document<'a> = { DocumentPath : string }


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
      PrintQuality: DocMakePrintQuality }



type State = 
    { MakeName: int -> string
      NameIndex: int }

let private incrNameIndex (st:State) : State = 
    let i = st.NameIndex in {st with NameIndex = i + 1 }


type BuildMonad<'res,'a> = 
    private BuildMonad of ((Env * 'res) -> StringBuilder -> State -> (State * Answer<'a>))

let inline private apply1 (ma : BuildMonad<'res,'a>) (handle:Env * 'res) (sbuf:StringBuilder) (st:State) : State *  Answer<'a> = 
    let (BuildMonad f) = ma in f handle sbuf st


let inline private unitM (x:'a) : BuildMonad<'res,'a> = 
    BuildMonad (fun _ _ st -> st, Ok x)


let inline private bindM (ma:BuildMonad<'res,'a>) (f : 'a -> BuildMonad<'res,'b>) : BuildMonad<'res,'b> =
    BuildMonad <| fun res sbuf st0 -> 
        let st1, ans = apply1 ma res sbuf st0
        match ans with
        | Err s -> (st1, Err s)
        | Ok a -> apply1 (f a) res sbuf st1


let inline private altM (ma:BuildMonad<'res,'a>) (mb:BuildMonad<'res,'a>) : BuildMonad<'res,'a> =
    BuildMonad <| fun res sbuf st0 -> 
        let sbuf2 = new StringBuilder() 
        let st1, ans = apply1 ma res sbuf2 st0
        match ans with 
        | Err s -> apply1 ma res sbuf st1
        | Ok a -> 
            sbuf.AppendLine (sbuf2.ToString()) |> ignore
            st1, Ok a

type BuildMonadBuilder() = 
    member self.Return x = unitM x
    member self.Bind (p,f) = bindM p f
    member self.Zero () = unitM ()

let (buildMonad:BuildMonadBuilder) = new BuildMonadBuilder()


// Common monadic operations
let fmapM (fn:'a -> 'b) (ma:BuildMonad<'res,'a>) : BuildMonad<'res,'b> = 
    BuildMonad <| fun res sbuf st0 ->
        let (st1,ans) =  apply1 ma res sbuf st0 
        match ans with
        | Err msg -> (st1, Err msg)
        | Ok a -> (st1, Ok <| fn a)


let mapM (fn:'a -> BuildMonad<'res,'b>) (xs:'a list) : BuildMonad<'res,'b list> =
    BuildMonad <| fun res sbuf state -> 
        let rec work ac (st0:State) ys = 
            match ys with
            | [] -> (st0, Ok <| List.rev ac)
            | z :: zs ->
                let (st1,ans) = apply1 (fn z) res sbuf st0
                match ans with
                | Err msg -> st1, Err msg 
                | Ok a -> work (a::ac) st1 zs
        work [] state xs

let forM (xs:'a list) (fn:'a -> BuildMonad<'res,'b>) : BuildMonad<'res,'b list> = 
    mapM fn xs


let mapMz (fn:'a -> BuildMonad<'res,'b>) (xs:'a list) : BuildMonad<'res,unit> =
    BuildMonad <| fun res sbuf state -> 
        let rec work st0 ys = 
            match ys with
            | [] -> (st0, Ok ())
            | z :: zs ->
                let (st1,ans) = apply1 (fn z) res sbuf st0
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
    BuildMonad <| fun res sbuf state -> 
        let rec work ac ix st0 ys = 
            match ys with
            | [] -> (st0, Ok <| List.rev ac)
            | z :: zs ->
                let (st1,ans) = apply1 (fn ix z) res sbuf st0
                match ans with
                | Err msg -> (st1, Err msg)
                | Ok a -> work (a::ac) (ix+1) st1 zs
        work [] 0 state xs

let mapiMz (fn:int -> 'a -> BuildMonad<'res,'b>) (xs:'a list) : BuildMonad<'res,unit> =
    BuildMonad <| fun res sbuf state -> 
        let rec work ix st0 ys = 
            match ys with
            | [] -> (st0, Ok ())
            | z :: zs ->
                let (st1,ans) =  apply1 (fn ix z) res sbuf st0
                match ans with
                | Err msg -> (st1, Err msg)
                | Ok _ -> work (ix+1) st1 zs
        work 0 state xs


// BuildMonad operations
let runBuildMonad (env:Env) (handle:'res) (stateZero:State) (ma:BuildMonad<'res,'a>) : State * string * Answer<'a>= 
    let sb = new StringBuilder () 
    match ma with 
    | BuildMonad fn -> let (s1,ans) = fn (env,handle) sb stateZero in (s1, sb.ToString(), ans)



/// Needs better name (launch has connotations of processes, threads, etc.)
let launch (handle:'res1) (ma:BuildMonad<'res1,'a>) : BuildMonad<'res2,'a> = 
    BuildMonad <| fun (env,_) sbuf st0 -> 
        let (st1,msg,ans) = runBuildMonad env handle st0 ma
        sbuf.AppendLine(msg) |> ignore
        (st1,ans)


let tell (msg:string) : BuildMonad<'res,unit> = 
    BuildMonad <| fun _ sbuf st0 -> 
        sbuf.Append(msg) |> ignore
        (st0, Ok ())


let tellLine (msg:string) : BuildMonad<'res,unit> = 
    BuildMonad <| fun _ sbuf st0 -> 
        sbuf.AppendLine(msg) |> ignore
        (st0, Ok ())


/// Note unit param to avoid value restriction.
let askU () : BuildMonad<'res,'res> = 
    BuildMonad <| fun (_,res) _ st0 -> (st0, Ok res)

let asksU (project:'res -> 'a) : BuildMonad<'res,'a> = 
    BuildMonad <| fun (_,res) _ st0 -> (st0, Ok (project res))

let localU (modify:'res -> 'res) (ma:BuildMonad<'res,'a>) : BuildMonad<'res,'a> = 
    BuildMonad <| fun (env,res) sbuf st0 -> apply1 ma (env, modify res) sbuf st0



/// The MakeName function is in State rather than Env, but we only provide an API
/// to run it within a context (cf. the Reader monad's local), rather than reset it 
/// imperatively.

let localState (modify:State -> State) (ma:BuildMonad<'res,'a>): BuildMonad<'res,'a> = 
    BuildMonad <| fun (env,res) sbuf st0 -> 
        let (_,ans) = apply1 ma (env,res) sbuf (modify st0)
        (st0,ans)

let withNameGen (namer:int -> string) (ma:BuildMonad<'res,'a>): BuildMonad<'res,'a> =  
    localState (fun s -> { s with  MakeName = namer; NameIndex = 1}) ma

let fileNameGen () : BuildMonad<'res, string> = 
    BuildMonad <| fun _ _  st0 -> 
        let i = st0.NameIndex
        (incrNameIndex st0, Ok <| st0.MakeName i)
