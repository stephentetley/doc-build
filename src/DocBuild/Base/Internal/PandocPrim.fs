// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Base.Internal


[<RequireQualifiedAccess>]
module PandocPrim = 
    
    open SLFormat.CommandOptions


    // One for SLFormat...
    let group (cmds:CmdOpt list) : CmdOpt = 
        List.fold (fun ac cmd -> ac ^+^ cmd) noArgument cmds
        

    /// pandoc --reference-doc="<customRef>" --from=markdown --to=docx+table_captions --standalone --output="<outputFile>" "<inputFile>"
    let outputDocxCommand (customRef:string option) 
                          (extraOpts: CmdOpt list)
                          (inputFile:string) 
                          (outputFile:string)  : CmdOpt list = 
        let customRef = 
            match customRef with
            | None -> noArgument
            | Some docx -> argument "--reference-doc" &= docx
        
        [ argument "--from"     &= "markdown"
        ; argument "--to"       &= "docx" &+ "table_captions"
        ; customRef
        ; group extraOpts
        ; argument "--standalone"
        ; argument "--output"   &= (doubleQuote outputFile)
        ; literal (argValue inputFile)
        ]
            
    /// Directly render Markdown to Pdf with Pandoc (and TeX).
    /// TeX must be installed.
    /// Default pdfengine is pdflatex
    /// pandoc --from=markdown --pdf-engine=<pdfEngine> --standalone --output=="<outputFile>" "<inputFile>"
    let outputPdfCommand (pdfEngine:string option) 
                         (extraOpts: CmdOpt list)
                         (inputFile:string) 
                         (outputFile:string) : CmdOpt list = 
        let pdfEngineOpt = 
            match pdfEngine with
            | None -> argument "--pdf-engine" &= "pdflatex"
            | Some engine -> argument "--pdf-engine" &= engine
        
        [ argument "--from"     &= "markdown"
        ; pdfEngineOpt
        ; argument "--standalone"
        ; argument "--output"   &= (doubleQuote outputFile)
        ; literal (argValue inputFile)
        ]
            


