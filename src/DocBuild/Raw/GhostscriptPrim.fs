// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Raw


[<RequireQualifiedAccess>]
module GhostscriptPrim = 
    
    open DocBuild.Base
    open DocBuild.Base.Shell
    open SLFormat
    open SLFormat.CommandOptions


    /// -dBATCH -dNOPAUSE -q -sDEVICE=pdfwrite [-dPDFSETTINGS=/screen]
    let private gsConcatOptions (quality:CmdOpt) : CmdOpt list =
        [ argument "-dBATCH"
        ; argument "-dNOPAUSE" 
        ; argument "-q" 
        ; argument "-sDEVICE" &= "pdfwrite"
        ; quality
        ]


    /// -sOutputFile="somefile.pdf"
    let private gsOutputFile (fileName:string) : CmdOpt = 
        argument "-sOutputFile" &= (doubleQuote fileName)
    
    /// "file1.pdf" "file2.pdf" ...
    let private gsInputFiles (fileNames:string list) : CmdOpt list  = 
        fileNames |> List.map (literal << doubleQuote)


    /// Apparently we cannot send multiline commands to execProcess.
    let concatCommand (quality:CmdOpt) 
                      (outputFile:string) 
                      (inputFiles: string list) : CmdOpt list = 
        gsConcatOptions quality @ [gsOutputFile outputFile] @  gsInputFiles inputFiles
        

