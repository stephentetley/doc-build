// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

#r "netstandard"
open System

// SLFormat & MarkdownDoc (not on nuget.org)
#I @"C:\Users\stephen\.nuget\packages\slformat\1.0.2-alpha-20190313\lib\netstandard2.0"
#r @"SLFormat.dll"
#I @"C:\Users\stephen\.nuget\packages\markdowndoc\1.0.1-alpha-20190314\lib\netstandard2.0"
#r @"MarkdownDoc.dll"

#load "..\src\DocBuild\Base\BaseDefinitions.fs"
#load "..\src\DocBuild\Base\FakeLikePrim.fs"
#load "..\src\DocBuild\Base\FilePaths.fs"
#load "..\src\DocBuild\Base\Common.fs"
#load "..\src\DocBuild\Base\Shell.fs"
#load "..\src\DocBuild\Base\DocMonad.fs"
#load "..\src\DocBuild\Base\DocMonadOperators.fs"
#load "..\src\DocBuild\Base\Document.fs"
#load "..\src\DocBuild\Base\Collection.fs"
#load "..\src\DocBuild\Base\FileOperations.fs"

open DocBuild.Base
open DocBuild.Base.DocMonad
open DocBuild.Base.DocMonadOperators


let WindowsEnv : DocBuildEnv = 
    let dataDir = DirectoryPath(System.IO.Path.Combine(__SOURCE_DIRECTORY__, "..", "data"))
    { WorkingDirectory = dataDir
      SourceDirectory = dataDir
      IncludeDirectory = DirectoryPath(dataDir <//> "include")
      GhostscriptExe = @"C:\programs\gs\gs9.15\bin\gswin64c.exe"
      PdftkExe = @"pdftk"
      PandocExe = @"pandoc" 
      PrintOrScreen = PrintQuality.Screen
      CustomStylesDocx = None
      }


let test01 () = 
    let sources = ["One.pdf"; "Two.pdf"; "Three.pdf"]
    let script = 
        docMonad { 
            let! docs = Collection.fromList <&&> mapM workingPdfDoc sources
            return docs
            }
    runDocMonadNoCleanup () WindowsEnv script
       
    
let test02 () = 
    test01 () 
        |> Result.map (fun col -> col.Elements)




let test04 () = 
    let sources = ["One.pdf"; "Two.pdf"]
    let script = 
        docMonad { 
            let! docs = 
                Collection.fromList <&&> mapM (workingPdfDoc) sources
            let! last = workingPdfDoc "Three.pdf"
            return (docs &^^ last)
            }
    runDocMonadNoCleanup () WindowsEnv script |> Result.map (fun col -> col.Elements)


let test05a () = 
    let script = assertM (mreturn false) "my error"
    runDocMonadNoCleanup () WindowsEnv script 

let test05b () = 
    let script = assertM (mreturn true) "my error"
    runDocMonadNoCleanup () WindowsEnv script 
