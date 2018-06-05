// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

module DocMake.Tasks.DocIntertitle

// Note - this is showing the limitations of using Fake (make) as the
// ``application style`` for building documents.
// An intertitle doesn't really deserve a build task, it is really just 
// an embellisher of other documents.



open Microsoft.Office.Interop

open DocMake.Base.OfficeUtils
open DocMake.Base.SimpleDocOutput

[<CLIMutable>]
type DocIntertitleParams = 
    { OutputFile: string
      DocumentCaption: string }

let DocIntertitleDefaults = 
    { OutputFile = @"intertitle.docx"
      DocumentCaption = "Title" }

    

let private addTitle  (caption:string) : DocOutput<unit> =
    tellStyledText Word.WdBuiltinStyle.wdStyleTitle (caption + "\n\n")
        



let DocIntertitle (setDocIntertitleParams: DocIntertitleParams -> DocIntertitleParams) : unit =
    let opts = DocIntertitleDefaults |> setDocIntertitleParams
    let procM = addTitle opts.DocumentCaption
    try 
        runDocOutput opts.OutputFile procM
    finally 
        ()
    




    



