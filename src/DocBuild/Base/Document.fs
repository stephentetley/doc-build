// Copyright (c) Stephen Tetley 2019
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


    /// Check the file exists and it's extension matches one of the supplied list.
    /// The path should be an absolute path.
    /// Throws an error within DocMonad on failure.
    let assertExistingFile (validFileExtensions:string list) 
                             (absPath:string) : DocMonad<'res,unit> = 
        if System.IO.File.Exists(absPath) then 
            let extension : string = System.IO.Path.GetExtension(absPath)
            let testExtension (ext:string) : bool = String.Equals(extension, ext, StringComparison.CurrentCultureIgnoreCase)
            if List.exists testExtension validFileExtensions then 
                dreturn ()
            else throwError <| sprintf "Not a %O file: '%s'" validFileExtensions absPath
        else throwError <| sprintf "Could not find file: '%s'" absPath 






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
                    (absPath:Uri) : DocMonad<'res,Document<'a>> = 
        docMonad { 
            do! assertExistingFile validFileExtensions absPath.LocalPath
            return Document(absPath)
            }



    /// Gets a Document from the working directory.
    let getWorkingDocument (validFileExtensions:string list) 
                           (relativeName:string) : DocMonad<'res,Document<'a>> = 
        docMonad { 
            let! path = askWorkingDirectory () |>> fun uri -> uri <//> relativeName
            return! getDocument validFileExtensions path 
            }


    /// Gets a Document from the source directory.
    let getSourceDocument (validFileExtensions:string list) 
                          (relativeName:string) : DocMonad<'res,Document<'a>> = 
        docMonad { 
            let! path = askSourceDirectory () |>> fun uri -> uri <//> relativeName
            return! getDocument validFileExtensions path 
            }


    /// Does not copies the source Document to working
    let getIncludeDocument (validFileExtensions:string list) 
                           (relativeName:string) : DocMonad<'res,Document<'a>> = 
       docMonad { 
            let! path = askIncludeDirectory () |>> fun uri -> uri <//> relativeName
            return! getDocument validFileExtensions path 
            }






    // ************************************************************************
    // Pdf file

    type PdfPhantom = class end

    type PdfFile = Document<PdfPhantom>

    /// Must have .pdf extension.
    let getPdfFile (absolutePath:string) : DocMonad<'res,PdfFile> = 
        new Uri(absolutePath) |> getDocument [".pdf"]

    /// Must have .pdf extension.
    let workingPdfFile (relativeName:string) : DocMonad<'res,PdfFile> = 
        getWorkingDocument [".pdf"] relativeName

    /// Must have .pdf extension.
    /// Writes a mutable copy to Working.
    let sourcePdfFile (relativeName:string) : DocMonad<'res,PdfFile> = 
        getSourceDocument [".pdf"] relativeName

    /// Must have .pdf extension.
    let includePdfFile (relativeName:string) : DocMonad<'res,PdfFile> = 
        getIncludeDocument [".pdf"] relativeName


    // ************************************************************************
    // Jpeg file

    type JpegPhantom = class end
    
    type JpegFile = Document<JpegPhantom>
    
    /// Must have .jpg or .jpeg extension.
    let getJpegFile (absolutePath:string) : DocMonad<'res,JpegFile> = 
        new Uri(absolutePath) |> getDocument [".jpg"; ".jpeg"]

    /// Must have .jpg or .jpeg extension.
    let workingJpegFile (relativeName:string) : DocMonad<'res,JpegFile> = 
        getWorkingDocument [".jpg"; ".jpeg"] relativeName


    /// Must have .jpg or .jpeg extension.
    /// Writes a mutable copy to Working.
    let sourceJpegFile (relativeName:string) : DocMonad<'res,JpegFile> = 
        getSourceDocument [".jpg"; ".jpeg"] relativeName
    
    /// Must have .jpg or .jpeg extension.
    let includeJpegFile (relativeName:string) : DocMonad<'res,JpegFile> = 
        getIncludeDocument [".jpg"; ".jpeg"] relativeName


    // ************************************************************************
    // Markdown file

    type MarkdownPhantom = class end

    type MarkdownFile = Document<MarkdownPhantom>

    /// Must have .md extension.
    let getMarkdownFile (absolutePath:string) : DocMonad<'res,MarkdownFile> = 
        new Uri(absolutePath) |> getDocument [".md"]

    /// Must have .md extension.
    let workingMarkdownFile (relativeName:string) : DocMonad<'res,MarkdownFile> = 
        getWorkingDocument [".md"] relativeName 

    /// Must have .md extension.
    /// Writes a mutable copy to Working.
    let sourceMarkdownFile (relativeName:string) : DocMonad<'res,MarkdownFile> = 
        getSourceDocument [".md"] relativeName 

    /// Must have .md extension.
    let includeMarkdownFile (relativeName:string) : DocMonad<'res,MarkdownFile> = 
        getIncludeDocument [".md"] relativeName 


    // ************************************************************************
    // Word file (.doc, .docx)

    type WordPhantom = class end

    type WordFile = Document<WordPhantom>

    /// Must have .doc or .docx extension.  
    let getWordFile (absolutePath:string) : DocMonad<'res,WordFile> = 
        new Uri(absolutePath) |> getDocument [".doc"; ".docx"]

    /// Must have .doc or .docx extension.    
    let workingWordFile (relativeName:string) : DocMonad<'res,WordFile> = 
        getWorkingDocument [".doc"; ".docx"] relativeName

    /// Must have .doc or .docx extension.
    /// Writes a mutable copy to Working.    
    let sourceWordFile (relativeName:string) : DocMonad<'res,WordFile> = 
        getSourceDocument [".doc"; ".docx"] relativeName

    /// Must have .doc or .docx extension.    
    let includeWordFile (relativeName:string) : DocMonad<'res,WordFile> = 
        getIncludeDocument [".doc"; ".docx"] relativeName


    // ************************************************************************
    // Excel file (.xls, .xlsx)

    type ExcelPhantom = class end

    type ExcelFile = Document<ExcelPhantom>

    /// Must have .xls or .xlsx or .xlsm extension.   
    let getExcelFile (absolutePath:string) : DocMonad<'res,ExcelFile> = 
        new Uri(absolutePath) |> getDocument [".xls"; ".xlsx"; ".xlsm"]

    /// Must have .xls or .xlsx or .xlsm extension. 
    let workingExcelFile (relativeName:string) : DocMonad<'res,ExcelFile> = 
        getWorkingDocument [".xls"; ".xlsx"; ".xlsm"] relativeName


    /// Must have .xls or .xlsx or .xlsm extension.
    /// Writes a mutable copy to Working.  
    let sourceExcelFile (relativeName:string) : DocMonad<'res,ExcelFile> = 
        getSourceDocument [".xls"; ".xlsx"; ".xlsm"] relativeName

    /// Must have .xls or .xlsx or .xlsm extension. 
    let includeExcelFile (relativeName:string) : DocMonad<'res,ExcelFile> = 
        getIncludeDocument [".xls"; ".xlsx"; ".xlsm"] relativeName


    // ************************************************************************
    // PowerPoint file (.ppt, .pptx)

    type PowerPointPhantom = class end

    type PowerPointFile = Document<PowerPointPhantom>

    /// Must have .ppt or .pptx extension. 
    let getPowerPointFile (absolutePath:string) : DocMonad<'res,PowerPointFile> = 
        new Uri(absolutePath) |> getDocument [".ppt"; ".pptx"]

    /// Must have .ppt or .pptx extension.
    let workingPowerPointFile (relativeName:string) : DocMonad<'res,PowerPointFile> = 
        getWorkingDocument [".ppt"; ".pptx"] relativeName 

    /// Must have .ppt or .pptx extension.
    /// Writes a mutable copy to Working.  
    let sourcePowerPointFile (relativeName:string) : DocMonad<'res,PowerPointFile> = 
        getSourceDocument [".ppt"; ".pptx"] relativeName 

    /// Must have .ppt or .pptx extension.
    let includePowerPointFile (relativeName:string) : DocMonad<'res,PowerPointFile> = 
        getIncludeDocument [".ppt"; ".pptx"] relativeName 


    // ************************************************************************
    // Text file (.txt)

    type TextPhantom = class end

    type TextFile = Document<TextPhantom>

    /// Must have .txt extension.  
    let getTextFile (absolutePath:string) : DocMonad<'res,TextFile> = 
        new Uri(absolutePath) |> getDocument[".txt"]


    /// Must have .txt extension.
    let workingTextFile (relativeName:string) : DocMonad<'res,TextFile> = 
        getWorkingDocument [".txt"] relativeName 

    /// Must have .txt extension.
    /// Writes a mutable copy to Working.  
    let sourceTextFile (relativeName:string) : DocMonad<'res,TextFile> = 
        getSourceDocument [".txt"] relativeName 

    /// Must have .txt extension.
    let includeTextFile (relativeName:string) : DocMonad<'res,TextFile> = 
        getIncludeDocument [".txt"] relativeName 
