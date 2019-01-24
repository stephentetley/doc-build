// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

#r "netstandard"

#load "..\src\DocBuild\Base\Common.fs"
#load "..\src\DocBuild\Base\Shell.fs"
#load "..\src\DocBuild\Base\DocMonad.fs"
#load "..\src\DocBuild\Base\DocMonadOperators.fs"
#load "..\src\DocBuild\Base\Document.fs"
#load "..\src\DocBuild\Base\Collection.fs"
#load "..\src\DocBuild\Base\FakeLike.fs"
#load "..\src\DocBuild\Base\FileIO.fs"

open DocBuild.Base
open DocBuild.Base.DocMonad
open DocBuild.Base.DocMonadOperators


let WindowsEnv : BuilderEnv = 
    let cwd = System.IO.Path.Combine(__SOURCE_DIRECTORY__, "..", "data")
    { WorkingDirectory = cwd
      GhostscriptExe = @"C:\programs\gs\gs9.15\bin\gswin64c.exe"
      PdftkExe = @"pdftk"
      PandocExe = @"pandoc"
      PandocReferenceDoc  = Some (cwd </> "custom-reference1.docx")
    }


let test01 () = 
    let sources = ["One.pdf"; "Two.pdf"; "Three.pdf"]
    let script = 
        docMonad { 
            let! docs = 
                makePdfCollection =<< mapM (askWorkingFile >=> getPdfFile) sources
            return docs
            }
    runDocMonadNoCleanup () WindowsEnv script
       
    
let test02 () = 
    test01 () 
        |> Result.map toList

let test03l () = 
    test01 () |> Result.map viewl

let test03r () = 
    let final ans = 
        match ans with
        | ViewR(col,one) -> printfn "col: %O" col; printfn "doc: %O" one
        | EmptyR -> printfn "EmptyR"
    test01 () |> Result.map (viewr >> final)

