// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Raw.Ghostscript


[<AutoOpen>]
module Ghostscript = 
    
    open DocBuild.Base
    open DocBuild.Base.Shell


    /// -dBATCH -dNOPAUSE -q -sDEVICE=pdfwrite [-dPDFSETTINGS=/screen]
    let private gsConcatOptions (quality:CommandArgs) : CommandArgs =
        noArg "-dBATCH" ^^ noArg "-dNOPAUSE" 
                        ^^ noArg "-q" 
                        ^^ reqArg "-sDEVICE" "pdfwrite"
                        ^^ quality


    /// -sOutputFile="somefile.pdf"
    let private gsOutputFile (fileName:string) : CommandArgs = 
        reqArg "-sOutputFile" (doubleQuote fileName)
    
    /// "file1.pdf" "file2.pdf" ...
    let private gsInputFiles (fileNames:string list) : CommandArgs = 
        fileNames |> List.map (noArg << doubleQuote) |> CommandArgs.Concat


    /// Apparently we cannot send multiline commands to execProcess.
    let makeGsConcatCommand (quality:CommandArgs) (outputFile:string) (inputFiles: string list) : CommandArgs = 
        gsConcatOptions quality  
            ^^ gsOutputFile outputFile 
            ^^ gsInputFiles inputFiles
        

