﻿// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Base

[<AutoOpen>]
module Document = 

    open System
    open System.Text.RegularExpressions
    open System.IO

    open DocBuild.Base.DocMonad
    open DocBuild.Base

    type FilePath = string

    /// The temp indicator is a suffix "Z0.." before the file extension
    let internal getNextTempName (filePath:FilePath) : FilePath =
        let root = System.IO.Path.GetDirectoryName filePath
        let justFile = Path.GetFileNameWithoutExtension filePath
        let extension  = System.IO.Path.GetExtension filePath

        let patt = @"Z(\d+)$"
        let result = Regex.Match(justFile, patt)
        let count = 
            if result.Success then 
                int <| result.Groups.Item(1).Value
            else 0
        let suffix = sprintf "Z%03d" (count+1)
        let newfile = sprintf "%s.%s%s" justFile suffix extension
        Path.Combine(root, newfile)

    let validateFile (fileExtensions:string list) (path:string) : DocMonad<string> = 
        if System.IO.File.Exists(path) then 
            let extension : string = System.IO.Path.GetExtension(path)
            let testExtension (ext:string) : bool = String.Equals(extension, ext, StringComparison.CurrentCultureIgnoreCase)
            if List.exists testExtension fileExtensions then 
                breturn path
            else throwError <| sprintf "Not a %s file: '%s'" (String.concat "," fileExtensions) path
        else throwError <| sprintf "Could not find file: '%s'" path  


    // ************************************************************************
    // Base Document type (to be wrapped)
    
    // TODO - maybe a Phantom type would be better then we aren't 
    // duplicating as much code.


    
    type Document<'a> = 
        | Document of FilePath 

        member x.Path 
            with get () : FilePath =
                match x with | Document(p) -> p

        /// ActiveFile is a mutable working copy of the original file.
        /// The original file is untouched.
        member x.NextTempName
            with get() : FilePath = 
                getNextTempName x.Path



    let getDocument (fileExtensions:string list) (filePath:string) : DocMonad<Document<'a>> = 
        docMonad { 
            let! path = validateFile fileExtensions filePath
            return Document(path)
            }

    
    // ************************************************************************
    // Pdf file

    type PdfPhantom = class end

    type PdfFile = Document<PdfPhantom>


    let getPdfFile (path:string) : DocMonad<PdfFile> = 
        getDocument [".pdf"] path 


    // ************************************************************************
    // Jpeg file

    type JpegPhantom = class end
    
    type JpegFile = Document<JpegPhantom>


    let getJpegFile (path:string) : DocMonad<JpegFile> = 
        getDocument [".jpg"; ".jpeg"] path

    // ************************************************************************
    // Markdown file

    type MarkdownPhantom = class end

    type MarkdownFile = Document<MarkdownPhantom>


    let getMarkdownFile (path:string) : DocMonad<MarkdownFile> = 
        getDocument [".md"] path 

    // ************************************************************************
    // Word file (.doc, .docx)

    type WordPhantom = class end

    type WordFile = Document<WordPhantom>


    let getWordFile (path:string) : DocMonad<WordFile> = 
        getDocument [".doc"; ".docx"] path



    // ************************************************************************
    // Excel file (.xls, .xlsx)

    type ExcelPhantom = class end

    type ExcelFile = Document<ExcelPhantom>


    /// Ignores .xlsm should it?
    let getExcelFile (path:string) : DocMonad<ExcelFile> = 
        getDocument [".xls"; ".xlsx"] path


    // ************************************************************************
    // PowerPoint file (.ppt, .pptx)

    type PowerPointPhantom = class end

    type PowerPointFile = Document<PowerPointPhantom>


    let getPowerPointFile (path:string) : DocMonad<PowerPointFile> = 
        getDocument [".ppt"; ".pptx"] path 

