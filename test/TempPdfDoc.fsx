// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

// Office deps
#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.Word"
#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Excel\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.Excel"
#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.PowerPoint\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.PowerPoint"
#I @"C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c"
#r "office"

#load "..\src\DocMake\Base\Common.fs"
#load "..\src\DocMake\Base\RunProcess.fs"
#load "..\src\DocMake\PdfDoc.fs"
#load "..\src\DocMake\WordDoc.fs"
open DocMake.PdfDoc
open DocMake.WordDoc

let demo01 () = 
    let working = System.IO.Path.Combine(__SOURCE_DIRECTORY__, "..", "data")
    let options = 
        { WorkingDirectory = working
        ; GhostscriptExe = @"C:\programs\gs\gs9.15\bin\gswin64c.exe" 
        ; PrintQuality = GsPdfScreen
        }
    let p1 = System.IO.Path.Combine(working, "One.pdf")
    let p2 = System.IO.Path.Combine(working, "Two.pdf")
    let p3 = System.IO.Path.Combine(working, "Three.pdf")
    let p4 = System.IO.Path.Combine(working, "FR-output2.docx")
    let d1 = (concat <| List.map pdfDoc [p1;p2;p3]) ^^ (wordDoc p4).ExportAsPdf(WordForScreen)
    d1.Save(options, "concat.pdf")


