// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Raw


[<RequireQualifiedAccess>]
module PandocPrim = 
    
    open SLFormat.CommandOptions

    open DocBuild.Base
    open DocBuild.Base.Shell

    
    /// pandoc --reference-doc="<customRef>" --from=markdown --to=docx+table_captions --standalone --output="<outputFile>" "<inputFile>"
    let outputDocxCommand (customRef:string option) 
                          (inputFile:string) 
                          (outputFile:string)  : CmdOpt list = 
        let customRef = 
            match customRef with
            | None -> noArgument
            | Some docx -> argument "--reference-doc" &= docx
        
        [ argument "--from"     &= "markdown"
        ; argument "--to"       &= "docx" &+ "table_captions"
        ; argument "--standalone"
        ; argument "--output"   &= outputFile
        ; literal (argValue inputFile)
        ; customRef
        ]
            




