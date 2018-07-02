// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

[<RequireQualifiedAccess>]
module DocMake.Tasks.MdToDoc



open DocMake.Base.Common
open DocMake.Base.OfficeUtils
open DocMake.Builder.BuildMonad
open DocMake.Builder.Document
open DocMake.Builder.Basis
open DocMake.Builder.ShellHooks

    


let makeCmd (infile:string) (outfile:string) : string = 
    sprintf "%s -f markdown -t docx -s -o %s" infile outfile 

let private docToPdfImpl (getHandle:'res-> PandocHandle) (mdDoc:MarkdownDoc) : BuildMonad<'res,WordDoc> =
    buildMonad { 
        let! outDoc = freshDocument () |>> documentChangeExtension "docx"
        let! _ =  pandocRunCommand getHandle <| makeCmd mdDoc.DocumentPath outDoc.DocumentPath
        return outDoc
    }

    
type MdToDocApi<'res> = 
    { MdToPdf : MarkdownDoc -> BuildMonad<'res, WordDoc> }

let makeAPI (getHandle:'res-> PandocHandle) : MdToDocApi<'res> = 
    { MdToPdf = docToPdfImpl getHandle }
