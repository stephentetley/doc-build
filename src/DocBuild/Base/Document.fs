// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Base

[<AutoOpen>]
module Document = 

    open System
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
                mreturn ()
            else throwError <| sprintf "Not a %O file: '%s'" validFileExtensions absPath
        else throwError <| sprintf "Could not find file: '%s'" absPath 






    // ************************************************************************
    // Base Document type 
    
    // We use a Phantom type parameter rather than a struct wrapper so 
    // we aren't duplicating the API for each Doc type.


    /// Work with System.Uri for file paths.
    type Document<'a> = 
        val private DocAbsPath: string
        val private DocTitle : string

        /// Title will be the file name, with directory information removed.
        new (absPath:string) = 
            { DocAbsPath = absPath; DocTitle = FileInfo(absPath).Name }
        
            
        new (absPath:string, title:string) = 
            { DocAbsPath = absPath; DocTitle = title }


        member x.Title 
            with get () : string = x.DocTitle

        member x.LocalPath
            with get () : string = FilePath(x.DocAbsPath).LocalPath

        member x.FileName
            with get () : string = 
                FileInfo(x.LocalPath).Name

                
    /// Set the document Title - title is the name of the document
    /// that might be used by some other process, e.g. to generate
    /// a table of contents.
    let setTitle (title:string) (doc:Document<'a>) : Document<'a> = 
        new Document<'a>(absPath=doc.LocalPath, title=title)


    /// Warning - this allows random access to the file system, not
    /// just the "Working"; "Include" and "Source" folders
    let getDocument (validFileExtensions:string list) 
                    (absPath:string) : DocMonad<'res,Document<'a>> = 
        docMonad { 
            do! assertExistingFile validFileExtensions absPath
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
            let! path = askIncludeDirectory () |>> fun dir -> dir <//> relativeName
            return! getDocument validFileExtensions path 
            }






    // ************************************************************************
    // Pdf file

    type PdfPhantom = class end

    type PdfDoc = Document<PdfPhantom>

    /// Must have .pdf extension.
    let getPdfDoc (absolutePath:string) : DocMonad<'res,PdfDoc> = 
        getDocument [".pdf"] absolutePath

    /// Must have .pdf extension.
    let workingPdfDoc (relativeName:string) : DocMonad<'res,PdfDoc> = 
        getWorkingDocument [".pdf"] relativeName

    /// Must have .pdf extension.
    /// Writes a mutable copy to Working.
    let sourcePdfDoc (relativeName:string) : DocMonad<'res,PdfDoc> = 
        getSourceDocument [".pdf"] relativeName

    /// Must have .pdf extension.
    let includePdfDoc (relativeName:string) : DocMonad<'res,PdfDoc> = 
        getIncludeDocument [".pdf"] relativeName


    // ************************************************************************
    // Jpeg file

    type JpegPhantom = class end
    
    type JpegDoc = Document<JpegPhantom>
    
    /// Must have .jpg or .jpeg extension.
    let getJpegDoc (absolutePath:string) : DocMonad<'res,JpegDoc> = 
        getDocument [".jpg"; ".jpeg"] absolutePath

    /// Must have .jpg or .jpeg extension.
    let workingJpegDoc (relativeName:string) : DocMonad<'res,JpegDoc> = 
        getWorkingDocument [".jpg"; ".jpeg"] relativeName


    /// Must have .jpg or .jpeg extension.
    /// Writes a mutable copy to Working.
    let sourceJpegDoc (relativeName:string) : DocMonad<'res,JpegDoc> = 
        getSourceDocument [".jpg"; ".jpeg"] relativeName
    
    /// Must have .jpg or .jpeg extension.
    let includeJpegDoc (relativeName:string) : DocMonad<'res,JpegDoc> = 
        getIncludeDocument [".jpg"; ".jpeg"] relativeName


    // ************************************************************************
    // Markdown file

    type MarkdownPhantom = class end

    type MarkdownDoc = Document<MarkdownPhantom>

    /// Must have .md extension.
    let getMarkdownDoc (absolutePath:string) : DocMonad<'res,MarkdownDoc> = 
        getDocument [".md"] absolutePath

    /// Must have .md extension.
    let workingMarkdownDoc (relativeName:string) : DocMonad<'res,MarkdownDoc> = 
        getWorkingDocument [".md"] relativeName 

    /// Must have .md extension.
    /// Writes a mutable copy to Working.
    let sourceMarkdownDoc (relativeName:string) : DocMonad<'res,MarkdownDoc> = 
        getSourceDocument [".md"] relativeName 

    /// Must have .md extension.
    let includeMarkdownDoc (relativeName:string) : DocMonad<'res,MarkdownDoc> = 
        getIncludeDocument [".md"] relativeName 


    // ************************************************************************
    // Word file (.doc, .docx)

    type WordPhantom = class end

    type WordDoc = Document<WordPhantom>

    /// Must have .doc or .docx extension.  
    let getWordDoc (absolutePath:string) : DocMonad<'res,WordDoc> = 
        getDocument [".doc"; ".docx"] absolutePath

    /// Must have .doc or .docx extension.    
    let workingWordDoc (relativeName:string) : DocMonad<'res,WordDoc> = 
        getWorkingDocument [".doc"; ".docx"] relativeName

    /// Must have .doc or .docx extension.
    /// Writes a mutable copy to Working.    
    let sourceWordDoc (relativeName:string) : DocMonad<'res,WordDoc> = 
        getSourceDocument [".doc"; ".docx"] relativeName

    /// Must have .doc or .docx extension.    
    let includeWordDoc (relativeName:string) : DocMonad<'res,WordDoc> = 
        getIncludeDocument [".doc"; ".docx"] relativeName


    // ************************************************************************
    // Excel file (.xls, .xlsx)

    type ExcelPhantom = class end

    type ExcelDoc = Document<ExcelPhantom>

    /// Must have .xls or .xlsx or .xlsm extension.   
    let getExcelDoc (absolutePath:string) : DocMonad<'res,ExcelDoc> = 
        getDocument [".xls"; ".xlsx"; ".xlsm"] absolutePath

    /// Must have .xls or .xlsx or .xlsm extension. 
    let workingExcelDoc (relativeName:string) : DocMonad<'res,ExcelDoc> = 
        getWorkingDocument [".xls"; ".xlsx"; ".xlsm"] relativeName


    /// Must have .xls or .xlsx or .xlsm extension.
    /// Writes a mutable copy to Working.  
    let sourceExcelDoc (relativeName:string) : DocMonad<'res,ExcelDoc> = 
        getSourceDocument [".xls"; ".xlsx"; ".xlsm"] relativeName

    /// Must have .xls or .xlsx or .xlsm extension. 
    let includeExcelDoc (relativeName:string) : DocMonad<'res,ExcelDoc> = 
        getIncludeDocument [".xls"; ".xlsx"; ".xlsm"] relativeName


    // ************************************************************************
    // PowerPoint file (.ppt, .pptx)

    type PowerPointPhantom = class end

    type PowerPointDoc = Document<PowerPointPhantom>

    /// Must have .ppt or .pptx extension. 
    let getPowerPointDoc (absolutePath:string) : DocMonad<'res,PowerPointDoc> = 
        getDocument [".ppt"; ".pptx"] absolutePath

    /// Must have .ppt or .pptx extension.
    let workingPowerPointDoc (relativeName:string) : DocMonad<'res,PowerPointDoc> = 
        getWorkingDocument [".ppt"; ".pptx"] relativeName 

    /// Must have .ppt or .pptx extension.
    /// Writes a mutable copy to Working.  
    let sourcePowerPointDoc (relativeName:string) : DocMonad<'res,PowerPointDoc> = 
        getSourceDocument [".ppt"; ".pptx"] relativeName 

    /// Must have .ppt or .pptx extension.
    let includePowerPointDoc (relativeName:string) : DocMonad<'res,PowerPointDoc> = 
        getIncludeDocument [".ppt"; ".pptx"] relativeName 


    // ************************************************************************
    // Text file (.txt)

    type TextPhantom = class end

    type TextDoc = Document<TextPhantom>

    /// Must have .txt extension.  
    let getTextDoc (absolutePath:string) : DocMonad<'res,TextDoc> = 
        getDocument[".txt"] absolutePath


    /// Must have .txt extension.
    let workingTextDoc (relativeName:string) : DocMonad<'res,TextDoc> = 
        getWorkingDocument [".txt"] relativeName 

    /// Must have .txt extension.
    /// Writes a mutable copy to Working.  
    let sourceTextDoc (relativeName:string) : DocMonad<'res,TextDoc> = 
        getSourceDocument [".txt"] relativeName 

    /// Must have .txt extension.
    let includeTextDoc (relativeName:string) : DocMonad<'res,TextDoc> = 
        getIncludeDocument [".txt"] relativeName 
