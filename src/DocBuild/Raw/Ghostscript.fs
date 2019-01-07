// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Raw


[<AutoOpen>]
module Ghostscript = 

    open DocBuild.Base.Common

    type GsPdfQuality = 
        | GsPdfScreen 
        | GsPdfEbook
        | GsPdfPrinter
        | GsPdfPrepress
        | GsPdfDefault
        | GsPdfNone
        member v.PdfQuality
            with get() = 
                match v with
                | GsPdfScreen ->  @"/screen"
                | GsPdfEbook -> @"/ebook"
                | GsPdfPrinter -> @"/printer"
                | GsPdfPrepress -> @"/prepress"
                | GsPdfDefault -> @"/default"
                | GsPdfNone -> ""


    type GhostscriptOptions = 
        { WorkingDirectory: string 
          GhostscriptExe: string 
          PrintQuality: GsPdfQuality
        }

 
    let private gsOptions (quality:GsPdfQuality) : string =
        match quality.PdfQuality with
        | "" -> @"-dBATCH -dNOPAUSE -q -sDEVICE=pdfwrite"
        | ss -> sprintf @"-dBATCH -dNOPAUSE -q -sDEVICE=pdfwrite -dPDFSETTINGS=%s" ss

    let private gsOutputFile (fileName:string) : string = 
        sprintf "-sOutputFile=\"%s\"" fileName

    let private gsInputFile (fileName:string) : string = sprintf "\"%s\"" fileName


    /// Apparently we cannot send multiline commands to execProcess.
    let makeGsConcatCommand (quality:GsPdfQuality) (outputFile:string) (inputFiles: string list) : string = 
        let line1 = gsOptions quality + " " + gsOutputFile outputFile
        let rest = List.map gsInputFile inputFiles
        String.concat " " (line1 :: rest)


    let runGhostscript (options:GhostscriptOptions) (command:string) : Choice<string,int> = 
        executeProcess options.WorkingDirectory options.GhostscriptExe command