// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause

#r "netstandard"
open System.IO
open System.Text.RegularExpressions

// Office deps
#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.Word"
#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Excel\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.Excel"
#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.PowerPoint\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.PowerPoint"
#I @"C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c"
#r "office"

// ImageMagick
#I @"C:\Users\stephen\.nuget\packages\Magick.NET-Q8-AnyCPU\7.11.1\lib\netstandard20"
#r @"Magick.NET-Q8-AnyCPU.dll"

// SLFormat & MarkdownDoc (not on nuget.org)
#I @"C:\Users\stephen\.nuget\packages\slformat\1.0.2-alpha-20190313\lib\netstandard2.0"
#r @"SLFormat.dll"
#I @"C:\Users\stephen\.nuget\packages\markdowndoc\1.0.1-alpha-20190314\lib\netstandard2.0"
#r @"MarkdownDoc.dll"

#load "..\src\DocBuild\Base\Internal\FakeLikePrim.fs"
#load "..\src\DocBuild\Base\Internal\FilePaths.fs"
#load "..\src\DocBuild\Base\Internal\Shell.fs"
#load "..\src\DocBuild\Base\Common.fs"
#load "..\src\DocBuild\Base\DocMonad.fs"
#load "..\src\DocBuild\Base\Document.fs"
#load "..\src\DocBuild\Base\Collection.fs"
#load "..\src\DocBuild\Base\FileOperations.fs"
#load "..\src\DocBuild\Raw\GhostscriptPrim.fs"
#load "..\src\DocBuild\Raw\PandocPrim.fs"
#load "..\src\DocBuild\Raw\PdftkPrim.fs"
#load "..\src\DocBuild\Raw\ImageMagickPrim.fs"
#load "..\src\DocBuild\Document\Pdf.fs"
#load "..\src\DocBuild\Document\Jpeg.fs"
#load "..\src\DocBuild\Document\Markdown.fs"
#load "..\src\DocBuild\Extra\PhotoBook.fs"

#load "..\src-msoffice\DocBuild\Office\Internal\Utils.fs"
#load "..\src-msoffice\DocBuild\Office\Internal\WordPrim.fs"
#load "..\src-msoffice\DocBuild\Office\Internal\ExcelPrim.fs"
#load "..\src-msoffice\DocBuild\Office\Internal\PowerPointPrim.fs"
#load "..\src-msoffice\DocBuild\Office\WordDocument.fs"
#load "..\src-msoffice\DocBuild\Office\ExcelDocument.fs"
#load "..\src-msoffice\DocBuild\Office\PowerPointDocument.fs"

open DocBuild.Base
open DocBuild.Document.Pdf
open DocBuild.Base.DocMonad



let temp01 () = 
    let fileName = "MyFile.Z001.jpg"
    printfn "%s" fileName

    let justFile = Path.GetFileNameWithoutExtension fileName
    printfn "%s" justFile
    
    let patt = @"Z(\d+)$"
    let result = Regex.Match(justFile, patt)
    if result.Success then 
        int <| result.Groups.Item(1).Value
    else
        0

/// The temp indicator is a suffix "Z0.." before the file extension
let getNextTempName (filePath:string) : string =
    let root = System.IO.Path.GetDirectoryName filePath
    let justFile = Path.GetFileNameWithoutExtension filePath
    let extension  = System.IO.Path.GetExtension filePath

    let patt = @"Z(\d+)$"
    let result = Regex.Match(justFile, patt)
    let count = 
        if result.Success then 
            int <| result.Groups.Item(1).Value
        else 0
    let suffix = sprintf "Z%03d" (count+1)
    let newfile = sprintf "%s.%s%s" justFile suffix extension
    Path.Combine(root, newfile)

let dataDump = """
InfoKey: CreationDate
InfoValue: D:20190110110137Z00'00'
PdfID0: 7b841f338439c24c13dbac0ec5f675b1
PdfID1: 7b841f338439c24c13dbac0ec5f675b1
NumberOfPages: 3
PageMediaBegin
PageMediaNumber: 1
PageMediaRotation: 0
PageMediaRect: 0 0 595.32 841.92
PageMediaDimensions: 595.32 841.92
PageMediaBegin
"""
 
let numPages() = 
    let patt = @"NumberOfPages: (\d+)"
    let result = Regex.Match(dataDump, patt)
    if result.Success then 
            result.Groups.Item(1).Value |> int |> Ok
    else 
        Error "numPages not found"


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

let makeResources (userRes:'userRes) : AppResources<'userRes> = 
    { GhostscriptExe = @"C:\programs\gs\gs9.15\bin\gswin64c.exe"
      PdftkExe = @"pdftk"
      PandocExe = @"pandoc" 
      UserResources = userRes
    }

let testCreateDir () =
    runDocMonadNoCleanup (makeResources ()) WindowsEnv <| 
        docMonad { 
            do! createWorkingSubdirectory @"TEMP_1\CHILD_A"
            return ()
        }

let traverse01 () =
    let operation (i:int) = 
        docMonad { 
            do printfn "%i" i
            return "ans"
            }
    runDocMonadNoCleanup (makeResources ()) WindowsEnv <| 
        mapMz operation [1;2;3;4]

let traverse01a () =
    let operation (i:int) = 
        if i < 3 then 
            docMonad { 
                do printfn "%i" i
                return "ans"
                }
        else throwError "large"
    runDocMonadNoCleanup (makeResources ()) WindowsEnv <| 
        mapMz operation [1;2;3;4]


let traverse02 () =
    let operation (i:int) = 
        docMonad { 
            do printfn "%i" i
            return i.ToString()
            }
    runDocMonadNoCleanup (makeResources ()) WindowsEnv <| 
        mapM operation [1;2;3;4]

let traverse02a () =
    let operation (i:int) = 
        if i < 3 then 
            docMonad { 
                do printfn "%i" i
                return i.ToString()
                }
        else throwError "large"
    runDocMonadNoCleanup (makeResources ()) WindowsEnv <| 
        mapM operation [1;2;3;4]

let traverse03 () =
    let operation (i:int) (s:string) = 
        docMonad { 
            do printfn "ix=%i val='%s'" i s
            return (i + int s)
            }
    runDocMonadNoCleanup (makeResources ()) WindowsEnv <| 
        mapiM operation ["1";"2";"3";"4"]

let traverse03a () =
    let operation (i:int) (s:string) = 
        if i < 3 then 
            docMonad { 
                do printfn "ix=%i val='%s'" i s
                return (i + int s)
                }
        else throwError "large"
    runDocMonadNoCleanup (makeResources ()) WindowsEnv <| 
        mapiM operation ["1";"2";"3";"4"]


type Env1 = 
    { NameField: string }

type UserResources = Map<string,obj>

let dynexperiment () = 
    let e:Env1 = { NameField = "Z001"}
    let d1 : UserResources = Map.empty
    let d2 = d1.Add("Env1", e :> obj)
    d2

let getEnv (resources:UserResources) : Env1 = 
    match Map.tryFind "Env1" resources with
    | Some o -> o :?> Env1
    | None -> failwith "Lookup"

let dummyLog () = 
    let logPath = Path.Combine(__SOURCE_DIRECTORY__, @"../data/log1.txt")
    use sw : StreamWriter = new StreamWriter(path = logPath)
    sw.WriteLine "Start"
    sw.WriteLine "End"


