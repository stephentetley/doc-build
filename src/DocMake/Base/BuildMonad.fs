module DocMake.Base.BuildMonad

open System.Text


type FailMsg = string

type Answer<'a> =
    | Err of FailMsg
    | Ok of 'a


type BuildMonad<'res,'a> = private BuildMonad of ('res -> StringBuilder -> Answer<'a>)

let inline private apply1 (ma : BuildMonad<'res,'a>) (handle:'res) (sbuf:StringBuilder) : Answer<'a> = 
    let (BuildMonad f) = ma in f handle sbuf

let inline private apply1Ex (ma : BuildMonad<'res,'a>) (handle:'res) (sbuf:StringBuilder) : 'a = 
    let (BuildMonad f) = ma
    match f handle sbuf with
    | Err msg -> failwith msg
    | Ok a -> a




let inline private unitM (x:'a) : BuildMonad<'res,'a> = BuildMonad (fun _ _ -> Ok x)


let inline private bindM (ma:BuildMonad<'res,'a>) (f : 'a -> BuildMonad<'res,'b>) : BuildMonad<'res,'b> =
    BuildMonad <| fun res sbuf -> 
        match apply1 ma res sbuf with
        | Err s -> Err s
        | Ok a -> apply1 (f a) res sbuf


let inline private altM (ma:BuildMonad<'res,'a>) (mb:BuildMonad<'res,'a>) : BuildMonad<'res,'a> =
    BuildMonad <| fun res sbuf -> 
        let sbuf2 = new StringBuilder() 
        match apply1 ma res sbuf2 with
        | Err s -> apply1 ma res sbuf
        | Ok a -> 
            sbuf.AppendLine (sbuf2.ToString()) |> ignore
            Ok a

type BuildMonadBuilder() = 
    member self.Return x = unitM x
    member self.Bind (p,f) = bindM p f
    member self.Zero () = unitM ()

let (buildMonad:BuildMonadBuilder) = new BuildMonadBuilder()


// Common monadic operations
let fmapM (fn:'a -> 'b) (ma:BuildMonad<'res,'a>) : BuildMonad<'res,'b> = 
    BuildMonad <| fun res sbuf ->
        match apply1 ma res sbuf with
        | Err msg -> Err msg
        | Ok a -> Ok (fn a)


let mapM (fn:'a -> BuildMonad<'res,'b>) (xs:'a list) : BuildMonad<'res,'b list> =
    BuildMonad <| fun res sbuf -> 
        let rec work ac ys = 
            match ys with
            | [] -> Ok <| List.rev ac
            | z :: zs ->
                match apply1 (fn z) res sbuf with
                | Err msg -> Err msg 
                | Ok a -> work (a::ac) zs
        work [] xs

let forM (xs:'a list) (fn:'a -> BuildMonad<'res,'b>) : BuildMonad<'res,'b list> = 
    mapM fn xs


let mapMz (fn:'a -> BuildMonad<'res,'b>) (xs:'a list) : BuildMonad<'res,unit> =
    BuildMonad <| fun res sbuf -> 
        let rec work ys = 
            match ys with
            | [] -> Ok ()
            | z :: zs ->
                match apply1 (fn z) res sbuf with
                | Err msg -> Err msg 
                | Ok _ -> work zs
        work xs


let forMz (xs:'a list) (fn:'a -> BuildMonad<'res,'b>) : BuildMonad<'res,unit> = 
    mapMz fn xs

let traverseM (fn: 'a -> BuildMonad<'res,'b>) (source:seq<'a>) : BuildMonad<'res,seq<'b>> = 
    BuildMonad <| fun res sbuf ->
        try 
            Seq.map (fun a -> apply1Ex (fn a) res sbuf) source |> Ok
        with 
        | ex -> Err (ex.Message)


let traverseMz (fn: 'a -> BuildMonad<'res,'b>) (source:seq<'a>) : BuildMonad<'res,unit> = 
    BuildMonad <| fun res sbuf ->
        try 
            Seq.fold (fun ac x -> 
                        let ans  = apply1Ex (fn x) res sbuf in ac) 
                     (Ok ())
                     source 
        with
        | ex -> Err (ex.Message)



let mapiM (fn:int -> 'a -> BuildMonad<'res,'b>) (xs:'a list) : BuildMonad<'res,'b list> = 
    BuildMonad <| fun res sbuf -> 
        let rec work ac ix ys = 
            match ys with
            | [] -> Ok <| List.rev ac
            | z :: zs ->
                match apply1 (fn ix z) res sbuf with
                | Err msg -> Err msg 
                | Ok a -> work (a::ac) (ix+1) zs
        work [] 0 xs

let mapiMz (fn:int -> 'a -> BuildMonad<'res,'b>) (xs:'a list) : BuildMonad<'res,unit> =
    BuildMonad <| fun res sbuf -> 
        let rec work ix ys = 
            match ys with
            | [] -> Ok ()
            | z :: zs ->
                match apply1 (fn ix z) res sbuf with
                | Err msg -> Err msg 
                | Ok _ -> work (ix+1) zs
        work 0 xs


// BuildMonad operations
let runBuildMonad (handle:'res) (ma:BuildMonad<'res,'a>) : string * Answer<'a>= 
    let sb = new StringBuilder () 
    match ma with 
    | BuildMonad fn -> let ans = fn handle sb in (sb.ToString(), ans)




let launch (handle:'res1) (ma:BuildMonad<'res1,'a>) : BuildMonad<'res2,'a> = 
    BuildMonad <| fun _ sbuf -> 
        let (s,ans) = runBuildMonad handle ma
        sbuf.AppendLine(s) |> ignore
        ans


let tell (msg:string) : BuildMonad<'res,unit> = 
    BuildMonad <| fun _ sbuf -> 
        sbuf.Append(msg) |> ignore
        Ok ()


let tellLine (msg:string) : BuildMonad<'res,unit> = 
    BuildMonad <| fun _ sbuf -> 
        sbuf.AppendLine(msg) |> ignore
        Ok ()


let temp01 () = 
    let sb1 = new System.Text.StringBuilder ()
    let sb2 = new System.Text.StringBuilder ()
    sb1.Append("hello") |> ignore
    sb2.Append(" world!") |> ignore
    sb1.Append(sb2.ToString()) |> ignore
    sb1.ToString()


