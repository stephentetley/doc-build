#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.Word"

#I @"C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c"
#r "office"

open Microsoft.Office.Interop

#load @"DocMake\Utils\Office.fs"
open DocMake.Utils.Office

#load @"DocMake\Tasks\DocPhotos.fs"
open DocMake.Tasks.DocPhotos


let relativeToProject (suffix:string) : string = 
    System.IO.Path.Combine(__SOURCE_DIRECTORY__, "..", suffix)

let test01 () = 
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


let test02 () = 
    let (opts: DocPhotosParams -> DocPhotosParams) = fun p -> 
        { p with 
            InputPaths = [ @"G:\work\photos1\TestFolder" ]
            OutputFile = relativeToProject @"data\photos1.docx"
            ShowFileName = true }
    DocPhotos opts



