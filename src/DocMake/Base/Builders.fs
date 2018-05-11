module DocMake.Base.Builders


// Open at .Interop rather than .Word then the Word API has to be qualified
open Microsoft.Office.Interop

open DocMake.Base.BuildMonad


type WordBuild<'a> = BuildMonad<Word.Application, 'a>
type WordPhantom = class end
type WordDoc = Document<WordPhantom>


type ExcelBuild<'a> = BuildMonad<Excel.Application, 'a>
type ExcelPhantom = class end
type ExcelDoc = Document<ExcelPhantom>


type PowerPointBuild<'a> = BuildMonad<PowerPoint.Application, 'a>
type PowerPointPhantom = class end
type PowerPointDoc = Document<PowerPointPhantom>
