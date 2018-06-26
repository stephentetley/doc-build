// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


module DocMake.Builder.WordHooks

open System.IO
open Microsoft.Office.Interop

open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis


type WordPhantom = class end
type WordDoc = Document<WordPhantom>


let private initWord () : Word.Application = 
    let app = new Word.ApplicationClass (Visible = true) :> Word.Application
    app

let private finalizeWord (app:Word.Application) : unit = app.Quit ()


let wordBuilderHook : BuilderHooks<Word.Application> = 
    { InitializeResource = initWord    
      FinalizeResource = finalizeWord }




let makeWordDoc (outputName:string) (proc:BuildMonad<'res, WordDoc>) :BuildMonad<'res, WordDoc> = 
    proc >>= renameTo outputName

let withDocxNamer (ma:BuildMonad<'res,'a>) : BuildMonad<'res,'a> = 
    withNameGen (sprintf "temp%03i.docx") ma