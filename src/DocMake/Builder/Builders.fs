// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


module DocMake.Builder.Builders

open System.IO
open Microsoft.Office.Interop

// Ideally don't use Fake
open Fake.Core


open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis



type WordBuild<'a> = BuildMonad<Word.Application, 'a>
type WordPhantom = class end
type WordDoc = Document<WordPhantom>

let makeWordDoc (outputName:string) (proc:BuildMonad<'res, WordDoc>) :BuildMonad<'res, WordDoc> = 
    proc >>= renameTo outputName



// TODO - this is not good as it can result in spawning many versions of Word.
// Instead we should ahve a single instance optionally in the 'res parameter.
let execWordBuild (ma:WordBuild<'a>) : BuildMonad<'res,'a> = 
    let app:Word.Application = new Word.ApplicationClass (Visible = true) :> Word.Application
    let namer:int -> string = fun i -> sprintf "temp%03i.docx" i
    withUserHandle app (fun (oApp:Word.Application) -> oApp.Quit()) (withNameGen namer ma)



type ExcelBuild<'a> = BuildMonad<Excel.Application, 'a>
type ExcelPhantom = class end
type ExcelDoc = Document<ExcelPhantom>

let execExcelBuild (ma:ExcelBuild<'a>) : BuildMonad<'res,'a> = 
    let app = new Excel.ApplicationClass(Visible = true) :> Excel.Application
    app.DisplayAlerts <- false
    app.EnableEvents <- false
    let namer:int -> string = fun i -> sprintf "temp%03i.docx" i
    let finalizer (oApp:Excel.Application) = 
        oApp.DisplayAlerts <- true
        oApp.EnableEvents <- true
        oApp.Quit ()
    withUserHandle app finalizer (withNameGen namer ma)



type PowerPointBuild<'a> = BuildMonad<PowerPoint.Application, 'a>
type PowerPointPhantom = class end
type PowerPointDoc = Document<PowerPointPhantom>

type PdfPhantom = class end
type PdfDoc = Document<PdfPhantom>


let makePdf (outputName:string) (proc:BuildMonad<'res, PdfDoc>) :BuildMonad<'res, PdfDoc> = 
    proc >>= renameTo outputName

// Shell helper:
let private shellRun toolPath (command:string)  (errMsg:string) : BuildMonad<'res, unit> = 
    let infoF (info:ProcStartInfo) =  
        { info with FileName =  toolPath; Arguments = command }
    buildMonad { 
        let i = Fake.Core.Process.execSimple infoF System.TimeSpan.MaxValue
        if (i = 0) then    
            return ()
        else
            do printfn "%s" toolPath
            do printfn "%s" command
            do! throwError errMsg
        }
            


// Ghostscript is Ghostscript (not camel case, i.e GhostScript)


// Having GsBuild/PdftkBuild is close to overkill as we will probably only
// ever run one-shot operations in them.
// However, having specialized types indicates that to the user that they 
// will need the respective applications/ toolsets installed.


type GsEnv = 
    { GhostscriptExePath: string }

type GsBuild<'a> = BuildMonad<GsEnv, 'a>


let execGsBuild (pathToGsExe:string) (ma:GsBuild<'a>) : BuildMonad<'res,'a> = 
    let gsEnv = { GhostscriptExePath = pathToGsExe }
    withUserHandle gsEnv (fun _ -> ()) ma


let gsRunCommand (command:string) : GsBuild<unit> = 
    buildMonad { 
        let! toolPath = asksU (fun (e:GsEnv) -> e.GhostscriptExePath) 
        do! shellRun toolPath command "GS failed"
    }

type PdftkEnv = 
    { PdftkExePath: string }

type PdftkBuild<'a> = BuildMonad<PdftkEnv, 'a>


let execPdftkBuild (pathToPdftkExe:string) (ma:PdftkBuild<'a>) : BuildMonad<'res,'a> = 
    let gsEnv = { PdftkExePath = pathToPdftkExe }
    withUserHandle gsEnv (fun _ -> ()) ma

let pdftkRunCommand (command:string) : PdftkBuild<unit> = 
    buildMonad { 
        let! toolPath = asksU (fun (e:PdftkEnv) -> e.PdftkExePath) 
        do! shellRun toolPath command "Pdftk failed"
    }