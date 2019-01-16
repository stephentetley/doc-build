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

let makeDoc (saiNumber:string) 
            (siteName:string)
            (logoPath:string) : Markdown = 
    concat [ logo logoPath
           ; nbsp2
           ; title1
           ; nbsp2
           ; title2 saiNumber siteName
           ]

let coversheet (saiNumber:string) 
               (siteName:string) 
               (logoPath:string)
               (outputFile:string) : DocMonad<'res,MarkdownFile> = 
    docMonad {         
        let markdown = makeDoc saiNumber siteName logoPath
        let! fullPath = localFile outputFile
        let! markdownFile = Markdown.saveMarkdown markdown fullPath
        return markdownFile
    } 

//let generateDocx (workingDirectory:string) (mdInputPath:string) (outputDocxName:string) : unit  =
//    let stylesDoc = @"include/custom-reference1.docx" 
//    runPandocDocx workingDirectory mdInputPath outputDocxName stylesDoc []