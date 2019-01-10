// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Raw


[<RequireQualifiedAccess>]
module PandocPrim = 

    open DocBuild.Base
    open DocBuild.Base.Shell
    
    /// pandoc --reference-doc="<customRef>" --from=markdown --to=docx+table_captions --standalone --output="<outputFile>" "<inputFile>"
    let outputDocxCommand (customRef:string option) 
                          (inputFile:string) 
                          (outputFile:string)  : CommandArgs = 
        let restArgs = 
            reqArg "--from" "markdown"
                ^^ reqArg "--to" "docx+table_captions"
                ^^ noArg "--standalone"
                ^^ reqArg "--output" (doubleQuote outputFile)
                ^^ noArg (doubleQuote inputFile)
        match customRef with
        | None -> restArgs 
        | Some docx -> reqArg "--reference-doc" (doubleQuote docx) ^^ restArgs
            




