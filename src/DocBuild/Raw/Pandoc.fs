// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Raw.Pandoc


[<AutoOpen>]
module Pandoc = 

    open DocBuild.Base.Common
    open DocBuild.Base.Shell
    
    /// pandoc --reference-doc="<customRef>" --from=markdown --to=docx+table_captions --standalone --output="<outputFile>" "<inputFile>"
    let makePandocDocxCommand (customRef:string) (inputFile:string) (outputFile:string)  : CommandArgs = 
        reqArg "--reference-doc" (doubleQuote customRef)
            ^^ reqArg "--from" "markdown"
            ^^ reqArg "--to" "docx+table_captions"
            ^^ noArg "--standalone"
            ^^ reqArg "--output" (doubleQuote outputFile)
            ^^ noArg (doubleQuote inputFile)




