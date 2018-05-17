#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.Word"
#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Excel\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.Excel"
#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.PowerPoint\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.PowerPoint"
#I @"C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c"
#r "office"

// FAKE is local to the project file
#I @"..\packages\FAKE.5.0.0-beta005\tools"
#r @"..\packages\FAKE.5.0.0-beta005\tools\FakeLib.dll"

#load @"DocMake\Base\Common.fs"
#load @"DocMake\Base\OfficeUtils.fs"
#load @"DocMake\Builder\BuildMonad.fs"
#load @"DocMake\Builder\Builders.fs"
#load @"DocMake\Builder\Basis.fs"
#load @"DocMake\Lib\DocFindReplace.fs"
open DocMake.Base.Common
open DocMake.Builder.BuildMonad
open DocMake.Builder.Builders
open DocMake.Lib.DocFindReplace


let matches1 = [ "#before", "after" ]


let temp01 () = 
    let env = 
        { WorkingDirectory = @"D:\coding\fsharp\DocMake\data"
          PrintQuality = DocMakePrintQuality.PqScreen }
    let proc = 
        buildMonad { 
            let! template = getTemplate @"D:\coding\fsharp\DocMake\data\findreplace1.docx"
            let! _ = docFindReplace matches1 template
            return ()
        }
    evalWordBuild env proc