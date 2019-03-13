// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Raw


[<RequireQualifiedAccess>]
module PandocPrim = 
    
    open SLFormat.CommandOptions

    open DocBuild.Base
    open DocBuild.Base.Shell

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
        ; group extraOpts
        ; argument "--standalone"
        ; argument "--output"   &= (doubleQuote outputFile)
        ; literal (argValue inputFile)
        ; customRef
        ]
            




