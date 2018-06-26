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

#load @"DocMake\Base\Common.fs"
#load @"DocMake\Base\FakeLike.fs"
#load @"DocMake\Base\OfficeUtils.fs"
#load @"DocMake\Builder\BuildMonad.fs"
#load @"DocMake\Builder\Basis.fs"
#load @"DocMake\Builder\WordHooks.fs"

#load @"DocMake\Tasks\DocFindReplace.fs"
#load @"DocMake\Tasks\DocToPdf.fs"
open DocMake.Base.Common
open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis
open DocMake.Builder.WordHooks
open DocMake.Tasks

let matches1 = [ "#before", "after" ]


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

