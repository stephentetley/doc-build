// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Raw


[<AutoOpen>]
module Pandoc = 

    

    type PandocOptions = 
        { WorkingDirectory: string 
          PandocExe: string 
          DocxReferenceDoc: string
        }

    // pandoc -f markdown -t docx+table_captions <INFILE> --reference-doc=<CUSTOM_REF> -s -o <OUTFILE>

    let makePandocCommand (inFile:string) (customRef:string) 
                                    (outFile:string) : string = 
        sprintf "-f markdown -t docx+table_captions \"%s\" --reference-doc=\"%s\" -s -o \"%s\""
                    inFile customRef outFile


