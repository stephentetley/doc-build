module DocMake.Builder.Builders

open System.IO
open Microsoft.Office.Interop

// Ideally don't use Fake
open Fake.Core.Process


open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis



type WordBuild<'a> = BuildMonad<Word.Application, 'a>
type WordPhantom = class end
type WordDoc = Document<WordPhantom>



let execWordBuild (ma:WordBuild<'a>) : BuildMonad<'res,'a> = 
    let app:Word.Application = new Word.ApplicationClass (Visible = true) :> Word.Application
    let namer:int -> string = fun i -> sprintf "temp%03i.docx" i
    withUserHandle app (fun (oApp:Word.Application) -> oApp.Quit()) (withNameGen namer ma)



type ExcelBuild<'a> = BuildMonad<Excel.Application, 'a>
type ExcelPhantom = class end
type ExcelDoc = Document<ExcelPhantom>


type PowerPointBuild<'a> = BuildMonad<PowerPoint.Application, 'a>
type PowerPointPhantom = class end
type PowerPointDoc = Document<PowerPointPhantom>

type PdfPhantom = class end
type PdfDoc = Document<PdfPhantom>


// Shell helper:
let private shellRun toolPath (command:string)  (errMsg:string) : BuildMonad<'res, unit> = 
    buildMonad { 
        let i = ExecProcess (fun info -> 
                    info.FileName <- toolPath
                    info.Arguments <- command) System.TimeSpan.MaxValue
        if (i = 0) then    
            return ()
        else
            do! throwError errMsg
        }
            


// Ghostscript is Ghostscript (not camel case, i.e GhostScript)


// Having GsBuild/PdftkBuild is close to overkill as will probably only
// ever run one-shot operations in them.
// However, having them indicates that I user will need to have the respective
// applications/ toolsets installed.


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