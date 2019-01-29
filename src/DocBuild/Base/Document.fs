// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Base

[<AutoOpen>]
module Document = 

    open System
    open System.Text.RegularExpressions
    open System.IO

    open DocBuild.Base.DocMonad


    /// The temp indicator is a suffix "Z0.." before the file extension
    let getNextTempName (filePath:string) : string =
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
        /// Pretend the suffix is an extension
        let prefix = Path.GetFileNameWithoutExtension justFile
        let newfile = sprintf "%s.%s%s" prefix suffix extension
        Path.Combine(root, newfile)

    let removeTempSuffix (filePath:string) : string =
        let root = System.IO.Path.GetDirectoryName filePath
        let justFile = Path.GetFileNameWithoutExtension filePath
        let extension  = System.IO.Path.GetExtension filePath

        let patt = @"Z(\d+)$"
        let result = Regex.Match(justFile, patt)
        if result.Success then 
            let prefix = Path.GetFileNameWithoutExtension justFile
            let newfile = Path.ChangeExtension(prefix, extension)
            Path.Combine(root, newfile)
        else filePath
        

    let validateFile (fileExtensions:string list) 
                     (path:string) : DocMonad<'res,string> = 
        if System.IO.File.Exists(path) then 
            let extension : string = System.IO.Path.GetExtension(path)
            let testExtension (ext:string) : bool = String.Equals(extension, ext, StringComparison.CurrentCultureIgnoreCase)
            if List.exists testExtension fileExtensions then 
                dreturn path
            else throwError <| sprintf "Not a %s file: '%s'" (String.concat "," fileExtensions) path
        else throwError <| sprintf "Could not find file: '%s'" path  


    // ************************************************************************
    // Base Document type 
    
    // We use a Phantom type parameter so we aren't duplicating 
    // API accessor wrappers for each Doc type.


    /// Work with System.Uri for file paths.
    type Document<'a> = 
        val DocPath : Uri
        val DocTitle : string

        /// Title will be the file name, with any temp information removed.
        new (path:string) = 
            { DocPath = new Uri(path); DocTitle = removeTempSuffix(path) }
        
        new (path:Uri) = 
            { DocPath = path; DocTitle = removeTempSuffix(path.AbsolutePath) }
            
        new (path:string, title:string) = 
            { DocPath = new Uri(path); DocTitle = title }

        new (path:Uri, title:string) = 
            { DocPath = path; DocTitle = title }

        member x.Path 
            with get () : Uri = x.DocPath

        member x.Title 
            with get () : string = x.DocTitle

        /// ActiveFile is a mutable working copy of the original file.
        /// The original file is untouched.
        member x.NextTempName
            with get() : Uri = 
                new Uri(getNextTempName <| x.DocPath.AbsolutePath)


    let getDocument (fileExtensions:string list) 
                    (filePath:string) : DocMonad<'res,Document<'a>> = 
        docMonad { 
            let! path = validateFile fileExtensions filePath
            return Document(path)
            }

    let getNamedDocument (fileExtensions:string list) 
                         (filePath:string) 
                         (title:string) : DocMonad<'res,Document<'a>> = 
        docMonad { 
            let! path = validateFile fileExtensions filePath
            return Document(path,title)
            }

    /// Set the document Title - title is the name of the document
    /// that might be used by some other process, e.g. to generate
    /// a table of contents.
    let setTitle (title:string) (doc:Document<'a>) : Document<'a> = 
        new Document<'a>(path=doc.Path, title=title)


    // ************************************************************************
    // Pdf file

    type PdfPhantom = class end

    type PdfFile = Document<PdfPhantom>

    /// Must have .pdf extension.
    let getPdfFile (path:string) : DocMonad<'res,PdfFile> = 
        getDocument [".pdf"] path 


    // ************************************************************************
    // Jpeg file

    type JpegPhantom = class end
    
    type JpegFile = Document<JpegPhantom>

    /// Must have .jpg .jpeg extension.
    let getJpegFile (path:string) : DocMonad<'res,JpegFile> = 
        getDocument [".jpg"; ".jpeg"] path

    // ************************************************************************
    // Markdown file

    type MarkdownPhantom = class end

    type MarkdownFile = Document<MarkdownPhantom>

    /// Must have .md extension.
    let getMarkdownFile (path:string) : DocMonad<'res,MarkdownFile> = 
        getDocument [".md"] path 

    // ************************************************************************
    // Word file (.doc, .docx)

    type WordPhantom = class end

    type WordFile = Document<WordPhantom>

    /// Must have .doc .docx extension.
    let getWordFile (path:string) : DocMonad<'res,WordFile> = 
        getDocument [".doc"; ".docx"] path



    // ************************************************************************
    // Excel file (.xls, .xlsx)

    type ExcelPhantom = class end

    type ExcelFile = Document<ExcelPhantom>


    /// Must have .xls .xlsx extension. Ignores .xlsm should it?
    let getExcelFile (path:string) : DocMonad<'res,ExcelFile> = 
        getDocument [".xls"; ".xlsx"] path


    // ************************************************************************
    // PowerPoint file (.ppt, .pptx)

    type PowerPointPhantom = class end

    type PowerPointFile = Document<PowerPointPhantom>

    /// Must have .ppt .pptx extension
    let getPowerPointFile (path:string) : DocMonad<'res,PowerPointFile> = 
        getDocument [".ppt"; ".pptx"] path 


    // ************************************************************************
    // Text file (.txt)

    type TextPhantom = class end

    type TextFile = Document<TextPhantom>

    /// Must have .txt extension
    let getTextFile (path:string) : DocMonad<'res,TextFile> = 
        getDocument [".txt"] path 


