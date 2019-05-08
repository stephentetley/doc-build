// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

#r "netstandard"
open System

// SLFormat & MarkdownDoc (not on nuget.org)
#I @"C:\Users\stephen\.nuget\packages\slformat\1.0.2-alpha-20190322\lib\netstandard2.0"
#r @"SLFormat.dll"
#I @"C:\Users\stephen\.nuget\packages\markdowndoc\1.0.1-alpha-20190508\lib\netstandard2.0"
#r @"MarkdownDoc.dll"

#load "..\src\DocBuild\Base\Internal\FakeLikePrim.fs"
#load "..\src\DocBuild\Base\Internal\FilePaths.fs"
#load "..\src\DocBuild\Base\Common.fs"
#load "..\src\DocBuild\Base\DocMonad.fs"
#load "..\src\DocBuild\Base\Document.fs"
#load "..\src\DocBuild\Base\Collection.fs"
#load "..\src\DocBuild\Base\FileOperations.fs"

open DocBuild.Base


let WindowsEnv : DocBuildEnv = 
    let dataDir = System.IO.Path.Combine(__SOURCE_DIRECTORY__, "..", "data")
    { WorkingDirectory = dataDir
      SourceDirectory = dataDir
      IncludeDirectories = [ dataDir </> "include" ]
      PrintOrScreen = PrintQuality.Screen
      PandocOpts = 
        { CustomStylesDocx = None
          PdfEngine = Some "pdflatex"
        }
      }

let makeResources (userRes:'res) : AppResources<'res> = 
    { GhostscriptExe = @"C:\programs\gs\gs9.15\bin\gswin64c.exe"
      PdftkExe = @"pdftk"
      PandocExe = @"pandoc" 
      UserResources = userRes
    }

let test01 () = 
    let sources = ["One.pdf"; "Two.pdf"; "Three.pdf"]
    let script = 
        docMonad { 
            let! docs = Collection.fromList <&&> mapM getWorkingPdfDoc sources
            return docs
            }
    runDocMonadNoCleanup (makeResources ()) WindowsEnv script
       
    
let test02 () = 
    test01 () 
        |> Result.map (fun col -> col.Elements)




let test04 () = 
    let sources = ["One.pdf"; "Two.pdf"]
    let script = 
        docMonad { 
            let! docs = 
                Collection.fromList <&&> mapM (getWorkingPdfDoc) sources
            let! last = getWorkingPdfDoc "Three.pdf"
            return (docs &^^ last)
            }
    runDocMonadNoCleanup (makeResources ()) WindowsEnv script |> Result.map (fun col -> col.Elements)


let test05a () = 
    let script = assertM "my error" (mreturn false) 
    runDocMonadNoCleanup (makeResources ()) WindowsEnv script 

let test05b () = 
    let script = assertM "my error" (mreturn true)
    runDocMonadNoCleanup (makeResources ()) WindowsEnv script 
