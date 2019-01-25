// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

// It would be nice and simple to use a prewritten Markdown
// doc and use find and replace.

module Coversheet

open MarkdownDoc

open DocBuild.Base
open DocBuild.Base.DocMonad
open DocBuild.Document


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


//type Item = 
//    { SaiNumber: string 
//      SiteName: string }

let nbsp2 : Markdown = 
    preformatted [nbsp; nbsp]

let doubleQuote (s:string) : string = sprintf "\"%s\"" s

let changeSlashes (path:string) : string = 
    path.Replace('\\', '/')


/// @"include/YW-logo.jpg"
let logo (includePath:string) : Markdown = 
    tile (inlineImage (text " ") (changeSlashes includePath) None)

let title1 : Markdown = 
    h1 (text "T0975 - Event Duration Monitoring Phase 2 (EDM2)")
    

let title2 (sai:string) (name:string) : Markdown = 
    h2 (text sai ^+^ text name)



let contents (workItems:string list) : Markdown = 
    h3 (text "Contents") ^@^ unordList (List.map (tile << text) workItems)

let documentControl : Markdown = 
    h3 (text "Document Control")


let columnSpecs : ColumnSpec list = 
    [  { Width = 20; Alignment = Alignment.AlignLeft }
    ;  { Width = 14; Alignment = Alignment.AlignLeft }
    ;  { Width = 16; Alignment = Alignment.AlignLeft }
    ;  { Width = 24; Alignment = Alignment.AlignLeft }
    ]
/// | 1.0 | S Tetley | 18/10/2018 | For EDMS |
let controlTable (author:string) : Markdown = 

    let nowstring = System.DateTime.Now.ToShortDateString()

    let makeHeaderCell (s:string) : Markdown = 
        text s |> doubleAsterisks |> tile

    let makeCell (s:string) : Markdown = text s |> tile

    let headers = 
        List.map makeHeaderCell ["Revision"; "Prepared By"; "Date"; "Comments"]
    let row1 = 
        List.map makeCell ["1.0"; author; nowstring; "For EDMS"]
    let row2 = [tile nbsp; tile empty; tile empty; tile empty]
    gridTable columnSpecs [headers;row1;row2] true


let makeDoc (saiNumber:string) 
            (siteName:string)
            (logoPath:string)
            (author:string) : Markdown = 
    concat [ logo logoPath
           ; nbsp2
           ; title1
           ; nbsp2
           ; title2 saiNumber siteName
           ; nbsp2
           ; documentControl 
           ; controlTable author
           ]

let coversheet (saiNumber:string) 
               (siteName:string) 
               (logoPath:string)
               (author:string)  
               (outputFile:string) : DocMonad<'res,MarkdownFile> = 
    docMonad {         
        let markdown = makeDoc saiNumber siteName logoPath author
        let! fullPath = changeToLocalFile outputFile
        let! markdownFile = Markdown.saveMarkdown fullPath markdown
        return markdownFile
    } 

//let generateDocx (workingDirectory:string) (mdInputPath:string) (outputDocxName:string) : unit  =
//    let stylesDoc = @"include/custom-reference1.docx" 
//    runPandocDocx workingDirectory mdInputPath outputDocxName stylesDoc []