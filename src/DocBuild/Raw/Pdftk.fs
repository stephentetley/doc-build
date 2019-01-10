// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Raw.Pdftk


// TODO - can use pdftk to get number of pages:
// > pdftk mydoc.pdf dump_data 
// Look for NumberOfPages in the output


[<AutoOpen>]
module Pdftk = 

    open System.Text.RegularExpressions

    open DocBuild.Base
    open DocBuild.Base.Shell

    

    /// Apparently we cannot send multiline commands to execProcess.
    let makePdftkDumpDataCommand (inputFile: string) : CommandArgs = 
        noArg (doubleQuote inputFile) ^^ noArg "dump_data"


    /// Seacrh for number of pages in a dump_data from Pdftk
    /// NumberOfPages: 3
    let regexSearchNumberOfPages (dumpData:string) : Result<int,ErrMsg> = 
        let patt = @"NumberOfPages: (\d+)"
        let result = Regex.Match(dumpData, patt)
        if result.Success then 
                result.Groups.Item(1).Value |> int |> Ok
        else 
            Error "regexSearchNumberOfPages 'NumberOfPages' not found"
        

