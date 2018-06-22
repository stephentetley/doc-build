// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

// FAKE is local to the project file
#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.Word"
#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Excel\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.Excel"
#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.PowerPoint\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.PowerPoint"
#I @"C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c"
#r "office"

open Microsoft.Office.Interop


// FAKE dependencies are getting onorous...
#I @"..\packages\FAKE.5.0.0-rc016.225\tools"
#r @"FakeLib.dll"
#I @"..\packages\Fake.Core.Globbing.5.0.0-beta021\lib\net46"
#r @"Fake.Core.Globbing.dll"
#I @"..\packages\Fake.IO.FileSystem.5.0.0-rc017.237\lib\net46"
#r @"Fake.IO.FileSystem.dll"
#I @"..\packages\Fake.Core.Trace.5.0.0-rc017.237\lib\net46"
#r @"Fake.Core.Trace.dll"
#I @"..\packages\Fake.Core.Process.5.0.0-rc017.237\lib\net46"
#r @"Fake.Core.Process.dll"



#load @"DocMake\Base\Common.fs"
#load @"DocMake\Base\OfficeUtils.fs"
#load @"DocMake\Base\SimpleDocOutput.fs"
open DocMake.Base.OfficeUtils

#load @"DocMake\Tasks\DocPhotos.fs"
open DocMake.Tasks.DocPhotos


let dummy01 () = 
    let app = new Word.ApplicationClass (Visible = true)
    try 
        let doc = app.Documents.Add()
        
        let paras = ["Page One";"Page Two";"Page Three"]
        List.iter (fun para -> 
                    let mutable rng = doc.GoTo(What = refobj Word.WdGoToItem.wdGoToBookmark, Name = refobj "\EndOfDoc")
                    rng.Text <- para 
                    rng <- doc.GoTo(What = refobj Word.WdGoToItem.wdGoToBookmark, Name = refobj "\EndOfDoc")
                    rng.InsertBreak(Type = refobj Word.WdBreakType.wdPageBreak) )
                paras
        doc.SaveAs(FileName= refobj @"E:\coding\fsharp\DocMake\data\TESTDOC1.docx")
        doc.Close(SaveChanges = refobj false)
        printfn "todo"
    finally 
        app.Quit ()


let test01 () = 
    let (opts: DocPhotosParams -> DocPhotosParams) = fun p -> 
        { p with 
            InputPaths = [ @"G:\work\photos1\TestFolder" ]
            OutputFile = @"G:\work\photos1\photodoc1.docx"
            ShowFileName = true 
            DocumentTitle = Some "Survey Photos" }
    DocPhotos opts



