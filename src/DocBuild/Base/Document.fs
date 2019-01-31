﻿// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Base

[<AutoOpen>]
module Document = 

    open System
    open System.Text.RegularExpressions
    open System.IO

    open DocBuild.Base
    open DocBuild.Base.DocMonad
    open DocBuild.Base.DocMonadOperators



    let validateExistingFile (validFileExtensions:string list) 
                             (path:Uri) : DocMonad<'res,Uri> = 
        if System.IO.File.Exists(path.LocalPath) then 
            let extension : string = System.IO.Path.GetExtension(path.LocalPath)
            let testExtension (ext:string) : bool = String.Equals(extension, ext, StringComparison.CurrentCultureIgnoreCase)
            if List.exists testExtension validFileExtensions then 
                dreturn path
            else throwError <| sprintf "Not a %O file: '%s'" validFileExtensions path.LocalPath
        else throwError <| sprintf "Could not find file: '%s'" path.LocalPath  






    // ************************************************************************
    // Base Document type 
    
    // We use a Phantom type parameter rather than a struct wrapper so 
    // we aren't duplicating the API for each Doc type.


    /// Work with System.Uri for file paths.
    type Document<'a> = 
        val private DocUri: Uri
        val private DocTitle : string

        /// Title will be the file name, with directory information removed.
        new (path:string) = 
            { DocUri = new Uri(path); DocTitle = FileInfo(path).Name }
        
        /// uri should be an absolute path.
        new (uri:Uri) = 
            { DocUri = uri; DocTitle = FileInfo(uri.LocalPath).Name }
            
        new (path:string, title:string) = 
            { DocUri = new Uri(path); DocTitle = title }

        /// uri should be an absolute path.            
        new (uri:Uri, title:string) = 
            { DocUri = uri; DocTitle = title }

        member x.Uri 
            with get () : Uri = x.DocUri

        member x.Title 
            with get () : string = x.DocTitle

        member x.LocalPath
            with get () : string = x.DocUri.LocalPath

        member x.FileName
            with get () : string = 
                FileInfo(x.DocUri.LocalPath).Name

                
    /// Set the document Title - title is the name of the document
    /// that might be used by some other process, e.g. to generate
    /// a table of contents.
    let setTitle (title:string) (doc:Document<'a>) : Document<'a> = 
        new Document<'a>(uri=doc.Uri, title=title)


    /// Warning - this allows random access to the file system, not
    /// just the "Working"; "Include" and "Source" folders
    let getDocument (validFileExtensions:string list) 
                    (filePath:Uri) : DocMonad<'res,Document<'a>> = 
        docMonad { 
            let! path = validateExistingFile validFileExtensions filePath
            return Document(path)
            }



    /// Gets a Document from the working directory.
    let getWorkingDocument (validFileExtensions:string list) 
                           (relativeName:string) : DocMonad<'res,Document<'a>> = 
        docMonad { 
            let! path = 
                askWorkingDirectory () |>> fun uri -> uri <//> relativeName
            let! uri = validateExistingFile validFileExtensions path
            return Document(uri)
            }


    /// Gets a Document from the source directory.
    let getSourceDocument (validFileExtensions:string list) 
                          (relativeName:string) : DocMonad<'res,Document<'a>> = 
        docMonad { 
            let! (path:Uri) = 
                askSourceDirectory () |>> fun uri -> uri <//> relativeName
            let! uri = validateExistingFile validFileExtensions path
            return Document(uri)
            }


    /// Does not copies the source Document to working
    let getIncludeDocument (validFileExtensions:string list) 
                           (relativeName:string) : DocMonad<'res,Document<'a>> = 
       docMonad { 
            let! (path:Uri) = 
                askIncludeDirectory () |>> fun uri -> uri <//> relativeName
            let! uri = validateExistingFile validFileExtensions path
            return Document(uri)
            }






    // ************************************************************************
    // Pdf file

    type PdfPhantom = class end

    type PdfFile = Document<PdfPhantom>

    /// Must have .pdf extension.
    let workingPdfFile (fileName:string) : DocMonad<'res,PdfFile> = 
        getWorkingDocument [".pdf"] fileName

    /// Must have .pdf extension.
    /// Writes a mutable copy to Working.
    let sourcePdfFile (fileName:string) : DocMonad<'res,PdfFile> = 
        getSourceDocument [".pdf"] fileName

    /// Must have .pdf extension.
    let includePdfFile (fileName:string) : DocMonad<'res,PdfFile> = 
        getIncludeDocument [".pdf"] fileName


    // ************************************************************************
    // Jpeg file

    type JpegPhantom = class end
    
    type JpegFile = Document<JpegPhantom>
    

    /// Must have .jpg or .jpeg extension.
    let workingJpegFile (fileName:string) : DocMonad<'res,JpegFile> = 
        getWorkingDocument [".jpg"; ".jpeg"] fileName


    /// Must have .jpg or .jpeg extension.
    /// Writes a mutable copy to Working.
    let sourceJpegFile (fileName:string) : DocMonad<'res,JpegFile> = 
        getSourceDocument [".jpg"; ".jpeg"] fileName
    
    /// Must have .jpg or .jpeg extension.
    let includeJpegFile (fileName:string) : DocMonad<'res,JpegFile> = 
        getIncludeDocument [".jpg"; ".jpeg"] fileName


    // ************************************************************************
    // Markdown file

    type MarkdownPhantom = class end

    type MarkdownFile = Document<MarkdownPhantom>

    /// Must have .md extension.
    let workingMarkdownFile (fileName:string) : DocMonad<'res,MarkdownFile> = 
        getWorkingDocument [".md"] fileName 

    /// Must have .md extension.
    /// Writes a mutable copy to Working.
    let sourceMarkdownFile (fileName:string) : DocMonad<'res,MarkdownFile> = 
        getSourceDocument [".md"] fileName 

    /// Must have .md extension.
    let includeMarkdownFile (fileName:string) : DocMonad<'res,MarkdownFile> = 
        getIncludeDocument [".md"] fileName 


    // ************************************************************************
    // Word file (.doc, .docx)

    type WordPhantom = class end

    type WordFile = Document<WordPhantom>

    /// Must have .doc or .docx extension.    
    let workingWordFile (fileName:string) : DocMonad<'res,WordFile> = 
        getWorkingDocument [".doc"; ".docx"] fileName

    /// Must have .doc or .docx extension.
    /// Writes a mutable copy to Working.    
    let sourceWordFile (fileName:string) : DocMonad<'res,WordFile> = 
        getSourceDocument [".doc"; ".docx"] fileName

    /// Must have .doc or .docx extension.    
    let includeWordFile (fileName:string) : DocMonad<'res,WordFile> = 
        getIncludeDocument [".doc"; ".docx"] fileName


    // ************************************************************************
    // Excel file (.xls, .xlsx)

    type ExcelPhantom = class end

    type ExcelFile = Document<ExcelPhantom>


    /// Must have .xls or .xlsx or .xlsm extension. 
    let workingExcelFile (fileName:string) : DocMonad<'res,ExcelFile> = 
        getWorkingDocument [".xls"; ".xlsx"; ".xlsm"] fileName


    /// Must have .xls or .xlsx or .xlsm extension.
    /// Writes a mutable copy to Working.  
    let sourceExcelFile (fileName:string) : DocMonad<'res,ExcelFile> = 
        getSourceDocument [".xls"; ".xlsx"; ".xlsm"] fileName

    /// Must have .xls or .xlsx or .xlsm extension. 
    let includeExcelFile (fileName:string) : DocMonad<'res,ExcelFile> = 
        getIncludeDocument [".xls"; ".xlsx"; ".xlsm"] fileName


    // ************************************************************************
    // PowerPoint file (.ppt, .pptx)

    type PowerPointPhantom = class end

    type PowerPointFile = Document<PowerPointPhantom>

    /// Must have .ppt or .pptx extension.
    let workingPowerPointFile (fileName:string) : DocMonad<'res,PowerPointFile> = 
        getWorkingDocument [".ppt"; ".pptx"] fileName 

    /// Must have .ppt or .pptx extension.
    /// Writes a mutable copy to Working.  
    let sourcePowerPointFile (fileName:string) : DocMonad<'res,PowerPointFile> = 
        getSourceDocument [".ppt"; ".pptx"] fileName 

    /// Must have .ppt or .pptx extension.
    let includePowerPointFile (fileName:string) : DocMonad<'res,PowerPointFile> = 
        getIncludeDocument [".ppt"; ".pptx"] fileName 


    // ************************************************************************
    // Text file (.txt)

    type TextPhantom = class end

    type TextFile = Document<TextPhantom>

    /// Must have .txt extension.
    let workingTextFile (fileName:string) : DocMonad<'res,TextFile> = 
        getWorkingDocument [".txt"] fileName 

    /// Must have .txt extension.
    /// Writes a mutable copy to Working.  
    let sourceTextFile (fileName:string) : DocMonad<'res,TextFile> = 
        getSourceDocument [".txt"] fileName 

    /// Must have .txt extension.
    let includeTextFile (fileName:string) : DocMonad<'res,TextFile> = 
        getIncludeDocument [".txt"] fileName 
