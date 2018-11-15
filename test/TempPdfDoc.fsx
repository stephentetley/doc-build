// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

#load "..\src\DocMake\Base\Common.fs"
#load "..\src\DocMake\Base\RunProcess.fs"
#load "..\src\DocMake\PdfDoc.fs"
open DocMake.Base.Common
open DocMake.PdfDoc

let demo01 () = 
    let working = System.IO.Path.Combine(__SOURCE_DIRECTORY__, "..", "data")
    let options = 
        { WorkingDirectory = working
        ; GhostscriptExe = @"C:\programs\gs\gs9.15\bin\gswin64c.exe" 
        ; PrintQuality = PdfPrintQuality.PdfScreen
        }
    let p1 = System.IO.Path.Combine(working, "One.pdf")
    let p2 = System.IO.Path.Combine(working, "Two.pdf")
    let p3 = System.IO.Path.Combine(working, "Three.pdf")
    let d1 = concat <| List.map pdfDoc [p1;p2;p3]
    d1.Save(options, "concat.pdf")


