module DocMake.Builder.Builders

open Microsoft.Office.Interop

open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis



type WordBuild<'a> = BuildMonad<Word.Application, 'a>
type WordPhantom = class end
type WordDoc = Document<WordPhantom>

let runWordBuild (env:Env) (ma:WordBuild<'a>) : State * string * Answer<'a>= 
    let st0 = 
        { MakeName = fun i -> sprintf "temp%03i.docx" i
          NameIndex = 1 }
    let app:Word.Application = new Word.ApplicationClass (Visible = true) :> Word.Application
    let st, bmlog, ans = runBuildMonad env app st0 ma
    app.Quit ()
    (st,bmlog,ans)

let evalWordBuild (env:Env) (ma:WordBuild<'a>) : 'a= 
    let st0 = 
        { MakeName = fun i -> sprintf "temp%03i.docx" i
          NameIndex = 1 }
    let app:Word.Application = new Word.ApplicationClass (Visible = true) :> Word.Application
    let finalizer = fun (handle:Word.Application) -> handle.Quit ()
    evalBuildMonad env app st0 finalizer ma
    

type ExcelBuild<'a> = BuildMonad<Excel.Application, 'a>
type ExcelPhantom = class end
type ExcelDoc = Document<ExcelPhantom>


type PowerPointBuild<'a> = BuildMonad<PowerPoint.Application, 'a>
type PowerPointPhantom = class end
type PowerPointDoc = Document<PowerPointPhantom>

type PdfPhantom = class end
type PdfDoc = Document<PdfPhantom>


// Ghostscript is Ghostscript (not camel case, i.e GhostScript)

type GsEnv = 
    { GhostscriptExePath: string }

type GsBuild<'a> = BuildMonad<GsEnv, 'a>

type PdftkEnv = 
    { PdftkExePath: string }

type PdftkBuild<'a> = BuildMonad<PdftkEnv, 'a>