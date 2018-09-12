// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.Word"
#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Excel\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.Excel"
#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.PowerPoint\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.PowerPoint"
#I @"C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c"
#r "office"

// FAKE is local to the project file
//#I @"..\packages\FAKE.5.0.0-beta005\tools"
//#r @"..\packages\FAKE.5.0.0-beta005\tools\FakeLib.dll"

#load "..\DocMake\DocMake\Base\Common.fs"
#load "..\DocMake\DocMake\Base\FakeLike.fs"
#load "..\DocMake\DocMake\Base\OfficeUtils.fs"
#load "..\DocMake\DocMake\Builder\BuildMonad.fs"
#load "..\DocMake\DocMake\Builder\Document.fs"
#load "..\DocMake\DocMake\Builder\Basis.fs"
#load "..\DocMake\DocMake\Tasks\DocFindReplace.fs"
#load "..\DocMake\DocMake\Tasks\DocToPdf.fs"
#load "..\DocMake\DocMake\WordBuilder.fs"
open DocMake.Builder.Document
open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis
open DocMake.Tasks
open DocMake.WordBuilder

let matches1 = [ "#before", "after" ]

// Out of date...

let test0 () = 
    let doc:WordDoc = { DocumentPath = @"D:\coding\fsharp\DocMake\data\TESTDOC1.docx"}
    printfn "%s" <| documentName doc
    printfn "%s" <| documentExtension doc
    printfn "%s" <| documentDirectory doc
    ()


let test01 () = 
    let env = 
        { WorkingDirectory = @"D:\coding\fsharp\DocMake\data"
          PrintQuality = DocMakePrintQuality.PqScreen
          PdfQuality = PdfPrintSetting.PdfPrint }
    let proc : WordBuild<unit> = 
        buildMonad { 
            let! template = getTemplateDoc @"D:\coding\fsharp\DocMake\data\findreplace1.docx"
            let! output = docFindReplace matches1 template
            let! a2 = renameTo @"findreplace2.docx" output 
            let! _ = docToPdf a2
            return ()
        }
    consoleRun env (execWordBuild proc)

