module DocMake.Base.Builders


// Open at .Interop rather than .Word then the Word API has to be qualified
open Microsoft.Office.Interop

open DocMake.Base.BuildMonad


type WordBuild<'a> = BuildMonad<Word.Application, 'a>
type WordPhantom = class end
type WordFile = Document<WordPhantom>


type ExcelBuild<'a> = BuildMonad<Excel.Application, 'a>
type ExcelPhantom = class end
type ExcelFile = Document<ExcelPhantom>


type PowerPointBuild<'a> = BuildMonad<PowerPoint.Application, 'a>
type PowerPointPhantom = class end
type PowerPointFile = Document<PowerPointPhantom>

type PdfPhantom = class end
type PdfFile = Document<PdfPhantom>


// Ghostscript is Ghostscript (not camel case, i.e GhostScript)

type GsEnv = 
    { GhostscriptExePath: string }

type GsBuild<'a> = BuildMonad<GsEnv, 'a>

type PdftkEnv = 
    { PdftkExePath: string }

type PdftkBuild<'a> = BuildMonad<PdftkEnv, 'a>