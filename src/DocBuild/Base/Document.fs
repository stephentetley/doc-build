// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Base

[<AutoOpen>]
module Document = 

    open System
    open System.IO

    open DocBuild.Base
    open DocBuild.Base.DocMonad



    /// Check the file exists and it's extension matches one of the supplied list.
    /// The path should be an absolute path.
    /// Throws an error within DocMonad on failure.
    let assertExistingFile (validFileExtensions:string list) 
                           (absPath:string) : DocMonad<'userRes,unit> = 
        if System.IO.File.Exists(absPath) then 
            let extension : string = System.IO.Path.GetExtension(absPath)
            let testExtension (ext:string) : bool = String.Equals(extension, ext, StringComparison.CurrentCultureIgnoreCase)
            if List.exists testExtension validFileExtensions then 
                mreturn ()
            else throwError <| sprintf "Not a %O file: '%s'" validFileExtensions absPath
        else throwError <| sprintf "Could not find file: '%s'" absPath 


    let isAbsolutePath (filePath:string) : bool = 
        try 
            match System.IO.Path.GetPathRoot(filePath) with
            | "" | null -> false
            | _ -> true
        with
        | _ -> false 





    // ************************************************************************
    // Base Document type 
    
    // We use a Phantom type parameter rather than a struct wrapper so 
    // we aren't duplicating the API for each Doc type.


    /// Work with string for file paths. System.Uri is unusable.
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

        member x.AbsolutePath
            with get () : string = x.DocAbsPath

        member x.FileName
            with get () : string = 
                FileInfo(x.DocAbsPath).Name

                
    /// Set the document Title - title is the name of the document
    /// that might be used by some other process, e.g. to generate
    /// a table of contents.
    let setTitle (title:string) (doc:Document<'a>) : Document<'a> = 
        new Document<'a>(absPath=doc.AbsolutePath, title=title)


    /// Warning - this allows random access to the file system, not
    /// just the "Working"; "Include" and "Source" folders
    let getDocument (validFileExtensions:string list) 
                    (absPath:string) : DocMonad<'userRes,Document<'a>> = 
        docMonad { 
            do! assertExistingFile validFileExtensions absPath
            return Document(absPath)
            }
        

    /// Gets a Document from the working directory.
    let getWorkingDocument (validFileExtensions:string list) 
                           (relativeName:string) : DocMonad<'userRes,Document<'a>> = 
        docMonad { 
            let! path = askWorkingDirectory () |>> fun dir -> dir </> relativeName
            return! getDocument validFileExtensions path 
            }


    /// Gets a Document from the source directory.
    let getSourceDocument (validFileExtensions:string list) 
                          (relativeName:string) : DocMonad<'userRes,Document<'a>> = 
        docMonad { 
            let! path = askSourceDirectory () |>> fun dir -> dir </> relativeName
            return! getDocument validFileExtensions path 
            }


    /// Does not copies the source Document to working
    let getIncludeDocument (validFileExtensions:string list) 
                           (relativeName:string) : DocMonad<'userRes,Document<'a>> = 
       docMonad { 
            let! (includeDirs :string list) = askIncludeDirectories () 
            let actions =
                List.map (fun dir -> getDocument validFileExtensions (dir </> relativeName)) includeDirs
            return! firstOfM actions
        }


    // ************************************************************************
    // Pdf file

    type PdfPhantom = class end

    type PdfDoc = Document<PdfPhantom>

    /// Must have .pdf extension.
    let getPdfDoc (absolutePath:string) : DocMonad<'userRes,PdfDoc> = 
        getDocument [".pdf"] absolutePath

    /// Must have .pdf extension.
    let workingPdfDoc (relativeName:string) : DocMonad<'userRes,PdfDoc> = 
        getWorkingDocument [".pdf"] relativeName

    /// Must have .pdf extension.
    /// Writes a mutable copy to Working.
    let sourcePdfDoc (relativeName:string) : DocMonad<'userRes,PdfDoc> = 
        getSourceDocument [".pdf"] relativeName

    /// Must have .pdf extension.
    let includePdfDoc (relativeName:string) : DocMonad<'userRes,PdfDoc> = 
        getIncludeDocument [".pdf"] relativeName


    // ************************************************************************
    // Jpeg file

    type JpegPhantom = class end
    
    type JpegDoc = Document<JpegPhantom>
    
    /// Must have .jpg or .jpeg extension.
    let getJpegDoc (absolutePath:string) : DocMonad<'userRes,JpegDoc> = 
        getDocument [".jpg"; ".jpeg"] absolutePath

    /// Must have .jpg or .jpeg extension.
    let workingJpegDoc (relativeName:string) : DocMonad<'userRes,JpegDoc> = 
        getWorkingDocument [".jpg"; ".jpeg"] relativeName


    /// Must have .jpg or .jpeg extension.
    /// Writes a mutable copy to Working.
    let sourceJpegDoc (relativeName:string) : DocMonad<'userRes,JpegDoc> = 
        getSourceDocument [".jpg"; ".jpeg"] relativeName
    
    /// Must have .jpg or .jpeg extension.
    let includeJpegDoc (relativeName:string) : DocMonad<'userRes,JpegDoc> = 
        getIncludeDocument [".jpg"; ".jpeg"] relativeName


    // ************************************************************************
    // Markdown file

    type MarkdownPhantom = class end

    type MarkdownDoc = Document<MarkdownPhantom>

    /// Must have .md extension.
    let getMarkdownDoc (absolutePath:string) : DocMonad<'userRes,MarkdownDoc> = 
        getDocument [".md"] absolutePath

    /// Must have .md extension.
    let workingMarkdownDoc (relativeName:string) : DocMonad<'userRes,MarkdownDoc> = 
        getWorkingDocument [".md"] relativeName 

    /// Must have .md extension.
    /// Writes a mutable copy to Working.
    let sourceMarkdownDoc (relativeName:string) : DocMonad<'userRes,MarkdownDoc> = 
        getSourceDocument [".md"] relativeName 

    /// Must have .md extension.
    let includeMarkdownDoc (relativeName:string) : DocMonad<'userRes,MarkdownDoc> = 
        getIncludeDocument [".md"] relativeName 


    // ************************************************************************
    // Word file (.doc, .docx)

    type WordPhantom = class end

    type WordDoc = Document<WordPhantom>

    /// Must have .doc or .docx extension.  
    let getWordDoc (absolutePath:string) : DocMonad<'userRes,WordDoc> = 
        getDocument [".doc"; ".docx"] absolutePath

    /// Must have .doc or .docx extension.    
    let workingWordDoc (relativeName:string) : DocMonad<'userRes,WordDoc> = 
        getWorkingDocument [".doc"; ".docx"] relativeName

    /// Must have .doc or .docx extension.  
    let sourceWordDoc (relativeName:string) : DocMonad<'userRes,WordDoc> = 
        getSourceDocument [".doc"; ".docx"] relativeName

    /// Must have .doc or .docx extension.    
    let includeWordDoc (relativeName:string) : DocMonad<'userRes,WordDoc> = 
        getIncludeDocument [".doc"; ".docx"] relativeName


    // ************************************************************************
    // Excel file (.xls, .xlsx)

    type ExcelPhantom = class end

    type ExcelDoc = Document<ExcelPhantom>

    /// Must have .xls or .xlsx or .xlsm extension.   
    let getExcelDoc (absolutePath:string) : DocMonad<'userRes,ExcelDoc> = 
        getDocument [".xls"; ".xlsx"; ".xlsm"] absolutePath

    /// Must have .xls or .xlsx or .xlsm extension. 
    let workingExcelDoc (relativeName:string) : DocMonad<'userRes,ExcelDoc> = 
        getWorkingDocument [".xls"; ".xlsx"; ".xlsm"] relativeName


    /// Must have .xls or .xlsx or .xlsm extension.
    /// Writes a mutable copy to Working.  
    let sourceExcelDoc (relativeName:string) : DocMonad<'userRes,ExcelDoc> = 
        getSourceDocument [".xls"; ".xlsx"; ".xlsm"] relativeName

    /// Must have .xls or .xlsx or .xlsm extension. 
    let includeExcelDoc (relativeName:string) : DocMonad<'userRes,ExcelDoc> = 
        getIncludeDocument [".xls"; ".xlsx"; ".xlsm"] relativeName


    // ************************************************************************
    // PowerPoint file (.ppt, .pptx)

    type PowerPointPhantom = class end

    type PowerPointDoc = Document<PowerPointPhantom>

    /// Must have .ppt or .pptx extension. 
    let getPowerPointDoc (absolutePath:string) : DocMonad<'userRes,PowerPointDoc> = 
        getDocument [".ppt"; ".pptx"] absolutePath

    /// Must have .ppt or .pptx extension.
    let workingPowerPointDoc (relativeName:string) : DocMonad<'userRes,PowerPointDoc> = 
        getWorkingDocument [".ppt"; ".pptx"] relativeName 

    /// Must have .ppt or .pptx extension.
    /// Writes a mutable copy to Working.  
    let sourcePowerPointDoc (relativeName:string) : DocMonad<'userRes,PowerPointDoc> = 
        getSourceDocument [".ppt"; ".pptx"] relativeName 

    /// Must have .ppt or .pptx extension.
    let includePowerPointDoc (relativeName:string) : DocMonad<'userRes,PowerPointDoc> = 
        getIncludeDocument [".ppt"; ".pptx"] relativeName 


    // ************************************************************************
    // Text file (.txt)

    type TextPhantom = class end

    type TextDoc = Document<TextPhantom>

    /// Must have .txt extension.  
    let getTextDoc (absolutePath:string) : DocMonad<'userRes,TextDoc> = 
        getDocument[".txt"] absolutePath


    /// Must have .txt extension.
    let workingTextDoc (relativeName:string) : DocMonad<'userRes,TextDoc> = 
        getWorkingDocument [".txt"] relativeName 

    /// Must have .txt extension.
    /// Writes a mutable copy to Working.  
    let sourceTextDoc (relativeName:string) : DocMonad<'userRes,TextDoc> = 
        getSourceDocument [".txt"] relativeName 

    /// Must have .txt extension.
    let includeTextDoc (relativeName:string) : DocMonad<'userRes,TextDoc> = 
        getIncludeDocument [".txt"] relativeName 
