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
    let dataDir = System.IO.Path.Combine(__SOURCE_DIRECTORY__, "..", "data")
    { WorkingDirectory = dataDir
      SourceDirectory = dataDir
      IncludeDirectory = dataDir </> "include"
      GhostscriptExe = @"C:\programs\gs\gs9.15\bin\gswin64c.exe"
      PdftkExe = @"pdftk"
      PandocExe = @"pandoc" }


let test01 () = 
    let sources = ["One.pdf"; "Two.pdf"; "Three.pdf"]
    let script = 
        docMonad { 
            let! docs = 
                Collection.fromList <&&> mapM (askWorkingFile >=> getPdfFile) sources
            return docs
            }
    runDocMonadNoCleanup () WindowsEnv script
       
    
let test02 () = 
    test01 () 
        |> Result.map Collection.toList

let test03l () = 
    test01 () |> Result.map Collection.viewl

let test03r () = 
    let final ans = 
        match ans with
        | Collection.ViewR(col,one) -> printfn "col: %O" col; printfn "doc: %O" one
        | Collection.EmptyR -> printfn "EmptyR"
    test01 () |> Result.map (Collection.viewr >> final)


let test04 () = 
    let sources = ["One.pdf"; "Two.pdf"]
    let script = 
        docMonad { 
            let! docs = 
                Collection.fromList <&&> mapM (askWorkingFile >=> getPdfFile) sources
            let! last = askWorkingFile "Three.pdf" >>= getPdfFile
            return (docs &>> last)
            }
    runDocMonadNoCleanup () WindowsEnv script |> Result.map Collection.toList