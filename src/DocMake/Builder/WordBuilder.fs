// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


module DocMake.Builder.WordBuilder

open System.IO
open Microsoft.Office.Interop



open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis


type WordPhantom = class end
type WordDoc = Document<WordPhantom>

let castToWordDoc (doc:Document<'a>) : WordDoc = castDocument doc


let private initWord () : Word.Application = 
    let app = new Word.ApplicationClass (Visible = true) :> Word.Application
    app

let private finalizeWord (app:Word.Application) : unit = app.Quit ()


let wordBuilderHook : BuilderHooks<Word.Application> = 
    { InitializeResource = initWord    
      FinalizeResource = finalizeWord }




let makeWordDoc (outputName:string) (proc:BuildMonad<'res, WordDoc>) :BuildMonad<'res, WordDoc> = 
    proc >>= renameTo outputName



//// TODO - Remove - run as a global single instance...
////
//// This is not good as it can result in spawning many versions of Word.
//// Instead we should ahve a single instance optionally in the 'res parameter.
//let execWordBuild (ma:WordBuild<'a>) : BuildMonad<'res,'a> = 
//    let app:Word.Application = new Word.ApplicationClass (Visible = true) :> Word.Application
//    let namer:int -> string = fun i -> sprintf "temp%03i.docx" i
//    withUserHandle app (fun (oApp:Word.Application) -> oApp.Quit()) (withNameGen namer ma)


    