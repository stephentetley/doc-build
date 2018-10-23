﻿// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

// Use FSharp.Data for CSV reading
#I @"..\packages\FSharp.Data.3.0.0-beta3\lib\net45"
#r @"FSharp.Data.dll"
open FSharp.Data

#I @"..\packages\__MY_LIBS__\lib\net45"
#r @"MarkdownDoc.dll"


open MarkdownDoc
open MarkdownDoc.Pandoc

let safeName (input:string) : string = 
    let parens = ['('; ')'; '['; ']'; '{'; '}']
    let bads = ['\\'; '/'; ':'; '?'] 
    let white = ['\n'; '\t']
    let ans1 = List.fold (fun (s:string) (c:char) -> s.Replace(c.ToString(), "")) input parens
    let ans2 = List.fold (fun (s:string) (c:char) -> s.Replace(c,'_')) ans1 bads
    let ans3 = List.fold (fun (s:string) (c:char) -> s.Replace(c,'_')) ans2 white
    ans3.Trim() 


// ****************************************************************************
// Build a document


type Item = 
    { Uid: string 
      Name: string
      Worklist: string list }

let nbsp2 : Markdown = 
    preformatted [nbsp; nbsp]

let logo : Markdown = 
    tile (inlineImage (rawtext " ") @"include/YW-logo.jpg" None)

let title1 : Markdown = 
    h1 (rawtext "T0975 - Event Duration Monitoring Phase 2 (EDM2)")
    

let title2 (sai:string) (name:string) : Markdown = 
    h2 (rawtext sai ^+^ rawtext name)



let contents (workItems:string list) : Markdown = 
    h3 (rawtext "Contents") + unordList (List.map (tile << rawtext) workItems)

let documentControl : Markdown = 
    h3 (rawtext "Document Control")

let makeDoc (item:Item) : Markdown = 
    concat [ logo
           ; nbsp2
           ; title1
           ; nbsp2
           ; title2 item.Uid item.Name
           ; nbsp2
           ; contents item.Worklist

           ]



let generateDocx (workingDirectory:string) (mdInputPath:string) (outputDocxName:string)  =
    let opts = 
        { ReferenceDoc = Some @"include/custom-reference1.docx" 
          DocxExtensions = extensions ["table_captions"] }
    runPandocDocx workingDirectory mdInputPath opts outputDocxName 

let test01 () = 
    printfn "%s" <| render 80 (makeDoc { Uid = "SAI01"; Name = "OTHER/SML"; Worklist = [] }) ;; 
