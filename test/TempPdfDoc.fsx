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

#I @"..\packages\__MY_LIBS__\lib\net45"
#r @"MarkdownDoc.dll"


#load "..\src\DocBuild\Internal\Common.fs"
#load "..\src\DocBuild\Internal\RunProcess.fs"
#load "..\src\DocBuild\PdfDoc.fs"
#load "..\src\DocBuild\ExcelDoc.fs"
#load "..\src\DocBuild\WordDoc.fs"
#load "..\src\DocBuild\PowerPointDoc.fs"
#load "..\src\DocBuild\MarkdownDoc.fs"
open DocBuild.PdfDoc
open DocBuild.ExcelDoc
open DocBuild.WordDoc
open DocBuild.PowerPointDoc
open DocBuild.MarkdownDoc

let getWorkingFile (name:string) = 
    let working = System.IO.Path.Combine(__SOURCE_DIRECTORY__, "..", "data")
    System.IO.Path.Combine(working, name)

let demo01 () = 
    let working = System.IO.Path.Combine(__SOURCE_DIRECTORY__, "..", "data")
    let gsOptions = 
        { WorkingDirectory = working
        ; GhostscriptExe = @"C:\programs\gs\gs9.15\bin\gswin64c.exe" 
        ; PrintQuality = GsPdfScreen
        }
    let pandocOptions = 
        { WorkingDirectory = working
        ; PandocExe = "pandoc"
        ; DocxReferenceDoc = @"include/custom-reference1.docx"
        }

    let p1 = getWorkingFile "One.pdf"
    let p2 = getWorkingFile "Two.pdf"
    let p3 = getWorkingFile "Three.pdf"
    let p4 = getWorkingFile "FR-output2.docx"
    let xl1 = excelDoc <| getWorkingFile "sheet1.xlsx"
    let md1 = markdownDoc <| getWorkingFile "sample.md"
    let d1 = (concat <| List.map pdfDoc [p1;p2;p3]) 
                ^^ (wordDoc p4).ExportAsPdf(WordForScreen)
                ^^ xl1.ExportAsPdf(true, ExcelQualityMinimum)
                ^^ md1.ExportAsPdf(pandocOptions)
    d1.Save(gsOptions, "concat.pdf")


