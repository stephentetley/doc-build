﻿// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

// It would be nice and simple to use a prewritten Markdown
// doc and use find and replace.

module Coversheet

open MarkdownDoc
open MarkdownDoc.Pandoc

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


let nbsp2 : Markdown = nbsp ^@^ nbsp

let doubleQuote (s:string) : string = sprintf "\"%s\"" s

let changeSlashes (path:string) : string = 
    path.Replace('\\', '/')


/// @"include/YW-logo.jpg"
let logo (includePath:string) : Markdown = 
    markdownText (inlineImage "" includePath None)

let title1 (titleText:string) : Markdown = h1 (text titleText)
    

let title2 (sai:string) (name:string) : Markdown = 
    h2 (text sai ^+^ text name)



let contents (workItems:string list) : Markdown = 
    h3 (text "Contents") ^@^ markdown (unorderedList (List.map (paraText << text) workItems))

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

    let makeHeaderCell (s:string) : ParaElement = 
        text s |> doubleAsterisks |> paraText

    let makeCell (s:string) : ParaElement = text s |> paraText

    let headers = 
        List.map makeHeaderCell ["Revision"; "Prepared By"; "Date"; "Comments"]
    let row1 = 
        List.map makeCell ["1.0"; author; nowstring; "For EDMS"]
    let row2 = [ParaElement.empty; ParaElement.empty; ParaElement.empty; ParaElement.empty]
    gridTable columnSpecs (Some headers) [row1; row2] 

type CoversheetConfig = 
    { LogoPath: string
      SaiNumber: string 
      SiteName: string  
      Author:string 
      Title: string }

let makeDoc (config:CoversheetConfig) : Markdown = 
    concatMarkdown
        <| [ logo config.LogoPath
           ; nbsp2
           ; title1 config.Title
           ; nbsp2
           ; title2 config.SaiNumber config.SiteName
           ; nbsp2
           ; documentControl 
           ; controlTable config.Author
           ]

let coversheet (config:CoversheetConfig)
               (outputRelName:string) : DocMonad<MarkdownDoc, 'res> = 
    docMonad {         
        let markdown = makeDoc config
        return! Markdown.saveMarkdown outputRelName markdown
    } 



