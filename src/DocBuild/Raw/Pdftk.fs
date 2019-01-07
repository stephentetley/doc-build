// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Raw


[<AutoOpen>]
module Pdftk = 

    open DocBuild.Base.Common

    
    type PdftkOptions = 
        { WorkingDirectory: string 
          PdftkExe: string 
        }

    let runPdftk (options:PdftkOptions) (command:string) : Choice<string,int> = 
        executeProcess options.WorkingDirectory options.PdftkExe command