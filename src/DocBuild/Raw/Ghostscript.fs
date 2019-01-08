// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Raw.Ghostscript


[<AutoOpen>]
module Ghostscript = 
    
    open DocBuild.Base.Common
    open DocBuild.Base.Shell

    type GsQuality = 
        | GsScreen 
        | GsEbook
        | GsPrinter
        | GsPrepress
        | GsDefault
        | GsNone
        member v.Quality
            with get() = 
                match v with
                | GsScreen ->  @"/screen"
                | GsEbook -> @"/ebook"
                | GsPrinter -> @"/printer"
                | GsPrepress -> @"/prepress"
                | GsDefault -> @"/default"
                | GsNone -> ""


    /// -dBATCH -dNOPAUSE -q -sDEVICE=pdfwrite [-dPDFSETTINGS=/screen]
    let private gsConcatOptions (quality:GsQuality) : CommandArgs =
        let common = 
            noArg "-dBATCH" ^^ noArg "-dNOPAUSE" ^^ noArg "-q" ^^ reqArg "-sDEVICE" "pdfwrite"
        match quality.Quality with
        | "" -> common
        | ss -> common ^^ reqArg "-dPDFSETTINGS" ss


    /// -sOutputFile="somefile.pdf"
    let private gsOutputFile (fileName:string) : CommandArgs = 
        reqArg "-sOutputFile" (doubleQuote fileName)
    
    /// "file1.pdf" "file2.pdf" ...
    let private gsInputFiles (fileNames:string list) : CommandArgs = 
        fileNames |> List.map (noArg << doubleQuote) |> CommandArgs.Concat


    /// Apparently we cannot send multiline commands to execProcess.
    let makeGsConcatCommand (quality:GsQuality) (outputFile:string) (inputFiles: string list) : CommandArgs = 
        gsConcatOptions quality  
            ^^ gsOutputFile outputFile 
            ^^ gsInputFiles inputFiles
        


    let runGhostscript (options:ProcessOptions) (command:CommandArgs) : ProcessResult = 
        executeProcess options command.Command