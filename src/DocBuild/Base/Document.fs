// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Base

[<AutoOpen>]
module Document = 

    open System
    open System.IO

    open DocBuild.Base
    open DocBuild.Base.Internal

    /// Check the file exists and it's extension matches one of the supplied list.
    /// The path should be an absolute path.
    /// Throws an error within DocMonad on failure.
    let assertExistingFile (validFileExtensions:string list) 
                           (absPath:string) : DocMonad<unit, 'userRes> = 
        if System.IO.File.Exists(absPath) then 
            let extension : string = System.IO.Path.GetExtension(absPath)
            let testExtension (ext:string) : bool = String.Equals(extension, ext, StringComparison.CurrentCultureIgnoreCase)
            if List.exists testExtension validFileExtensions then 
                mreturn ()
            else docError <| sprintf "Not a %O file: '%s'" validFileExtensions absPath
        else docError <| sprintf "Could not find file: '%s'" absPath 


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
    //type DocBody = 
    //    { DocAbsPath: string
    //      DocTitle: string }

    
    type Document<'a> = 
        val private DocAbsPath : string option
        val private DocTitle : string                        
        
        internal new (absPath:string option, title:string) = 
            { DocAbsPath = absPath; DocTitle = title }

        /// Title will be the file name, with directory information removed.
        new (absPath:string) = 
            { DocAbsPath = Some absPath; DocTitle = FileInfo(absPath).Name }
        
        new (absPath:string, title:string) = 
            { DocAbsPath = Some absPath; DocTitle = title }

        static member empty = 
            new Document<'a> (absPath = None, title = "empty")


        member x.Title 
            with get () : string = x.DocTitle

        member x.AbsolutePath
            with get () : string option = x.DocAbsPath

        member x.FileName
            with get () : string option = 
                x.AbsolutePath |> Option.map (fun path -> FileInfo(path).Name)

        member x.Extension
            with get () : string option = 
                x.AbsolutePath |> Option.map (fun path -> FileInfo(path).Extension)

                
    /// Set the document Title - title is the name of the document
    /// that might be used by some other process, e.g. to generate
    /// a table of contents.
    let setTitle (title:string) (doc:Document<'a>) : Document<'a> = 
        new Document<'a>(absPath=doc.AbsolutePath, title=title)


    let private isWorkingPath (absPath:string) : DocMonad<bool, 'userRes> = 
        askWorkingDirectory () |>> fun dir -> FilePaths.rootIsPrefix dir absPath


    let isWorkingDocument (doc : Document<'a>) : DocMonad<bool, 'userRes> = 
        match doc.AbsolutePath with
        | Some path -> isWorkingPath path
        | None -> mreturn false

    let getDocumentPath (doc : Document<'a>) : DocMonad<string, 'userRes> = 
        match doc.AbsolutePath with
        | Some path -> mreturn path
        | None -> docError "Invalid Document"

    let getDocumentFileName (doc : Document<'a>) : DocMonad<string, 'userRes> = 
        match doc.FileName with
        | Some name -> mreturn name
        | None -> docError "Invalid Document"



    let renameDocument (relativeName:string) 
                       (doc:Document<'a>) : DocMonad<Document<'a>, 'userRes> = 
        docMonad { 
            match! isWorkingDocument doc with
            | false -> return! docError "Rename failed document not in working directory."
            | true -> 
                let title = doc.Title
                let extension = doc.Extension
                let! absPath = getDocumentPath doc
                let directory = Path.GetDirectoryName(absPath)
                let dest = Path.Combine(directory, relativeName)
                do! liftOperation "IO error - File.Move" (fun _ -> File.Move(sourceFileName = absPath, destFileName = dest))
                return Document(absPath = dest, title = title)
        }


    let mandatory (docbuild : DocMonad<Document<'a>, 'userRes>) : DocMonad<Document<'a>, 'userRes> = 
        docMonad { 
            let! doc = docbuild
            match doc.AbsolutePath with
            | None -> return! docError "mandatory"
            | Some _ -> return doc
        }


    let nonMandatory (docbuild : DocMonad<Document<'a>, 'userRes>) : DocMonad<Document<'a>, 'userRes> = 
       docbuild <|> mreturn Document.empty


    /// Warning - this allows random access to the file system, not
    /// just the "Working"; "Include" and "Source" folders
    /// Condition - the document must exist.
    let getDocument (validFileExtensions:string list) 
                    (absPath:string) : DocMonad<Document<'a>, 'userRes> = 
        docMonad { 
            do! assertExistingFile validFileExtensions absPath
            return Document(absPath = absPath)
            }
        

    /// Gets a Document from the working directory.
    /// Condition - the document must exist.
    let getWorkingDocument (validFileExtensions:string list) 
                           (relativeName:string) : DocMonad<Document<'a>, 'userRes> = 
        docMonad { 
            let! path = askWorkingDirectory () |>> fun dir -> dir </> relativeName
            return! getDocument validFileExtensions path 
            }


    /// Gets a Document from the source directory.
    /// Condition - the document must exist.
    let getSourceDocument (validFileExtensions:string list) 
                          (relativeName:string) : DocMonad<Document<'a>, 'userRes> = 
        docMonad { 
            let! path = askSourceDirectory () |>> fun dir -> dir </> relativeName
            return! getDocument validFileExtensions path 
            }

    /// Gets a Document from the includes directories (searches directoties).
    /// Does not copies the source Document to working
    /// Condition - the document must exist.
    let getIncludeDocument (validFileExtensions:string list) 
                           (relativeName:string) : DocMonad<Document<'a>, 'userRes> = 
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
    /// Condition - the document must exist.
    let getPdfDoc (absolutePath:string) : DocMonad<PdfDoc, 'userRes> = 
        getDocument [".pdf"] absolutePath

    /// Must have .pdf extension.
    /// Condition - the document must exist.
    let getWorkingPdfDoc (relativeName:string) : DocMonad<PdfDoc, 'userRes> = 
        getWorkingDocument [".pdf"] relativeName

    /// Must have .pdf extension.
    /// Condition - the document must exist.
    let getSourcePdfDoc (relativeName:string) : DocMonad<PdfDoc, 'userRes> = 
        getSourceDocument [".pdf"] relativeName

    /// Must have .pdf extension.
    /// Condition - the document must exist.
    let getIncludePdfDoc (relativeName:string) : DocMonad<PdfDoc, 'userRes> = 
        getIncludeDocument [".pdf"] relativeName


    // ************************************************************************
    // Jpeg file

    type JpegPhantom = class end
    
    type JpegDoc = Document<JpegPhantom>
    
    /// Must have .jpg or .jpeg extension.
    /// Condition - the document must exist.
    let getJpegDoc (absolutePath:string) : DocMonad<JpegDoc, 'userRes> = 
        getDocument [".jpg"; ".jpeg"] absolutePath

    /// Must have .jpg or .jpeg extension.
    /// Condition - the document must exist.
    let getWorkingJpegDoc (relativeName:string) : DocMonad<JpegDoc, 'userRes> = 
        getWorkingDocument [".jpg"; ".jpeg"] relativeName


    /// Must have .jpg or .jpeg extension.
    /// Condition - the document must exist.
    let getSourceJpegDoc (relativeName:string) : DocMonad<JpegDoc, 'userRes> = 
        getSourceDocument [".jpg"; ".jpeg"] relativeName
    
    /// Must have .jpg or .jpeg extension.
    /// Condition - the document must exist.
    let getIncludeJpegDoc (relativeName:string) : DocMonad<JpegDoc, 'userRes> = 
        getIncludeDocument [".jpg"; ".jpeg"] relativeName


    // ************************************************************************
    // Markdown file

    type MarkdownPhantom = class end

    type MarkdownDoc = Document<MarkdownPhantom>

    /// Must have .md extension.
    /// Condition - the document must exist.
    let getMarkdownDoc (absolutePath:string) : DocMonad<MarkdownDoc, 'userRes> = 
        getDocument [".md"] absolutePath

    /// Must have .md extension.
    /// Condition - the document must exist.
    let getWorkingMarkdownDoc (relativeName:string) : DocMonad<MarkdownDoc, 'userRes> = 
        getWorkingDocument [".md"] relativeName 

    /// Must have .md extension.
    /// Condition - the document must exist.
    let getSourceMarkdownDoc (relativeName:string) : DocMonad<MarkdownDoc, 'userRes> = 
        getSourceDocument [".md"] relativeName 

    /// Must have .md extension.
    /// Condition - the document must exist.
    let getIncludeMarkdownDoc (relativeName:string) : DocMonad<MarkdownDoc, 'userRes> = 
        getIncludeDocument [".md"] relativeName 


    // ************************************************************************
    // Word file (.doc, .docx)

    type WordPhantom = class end

    type WordDoc = Document<WordPhantom>

    /// Must have .doc or .docx extension.  
    /// Condition - the document must exist.
    let getWordDoc (absolutePath:string) : DocMonad<WordDoc, 'userRes> = 
        getDocument [".doc"; ".docx"] absolutePath

    /// Must have .doc or .docx extension.
    /// Condition - the document must exist.
    let getWorkingWordDoc (relativeName:string) : DocMonad<WordDoc, 'userRes> = 
        getWorkingDocument [".doc"; ".docx"] relativeName

    /// Must have .doc or .docx extension.
    /// Condition - the document must exist.    
    let getSourceWordDoc (relativeName:string) : DocMonad<WordDoc, 'userRes> = 
        getSourceDocument [".doc"; ".docx"] relativeName

    /// Must have .doc or .docx extension.
    /// Condition - the document must exist.
    let getIncludeWordDoc (relativeName:string) : DocMonad<WordDoc, 'userRes> = 
        getIncludeDocument [".doc"; ".docx"] relativeName


    // ************************************************************************
    // Excel file (.xls, .xlsx)

    type ExcelPhantom = class end

    type ExcelDoc = Document<ExcelPhantom>

    /// Must have .xls or .xlsx or .xlsm extension. 
    /// Condition - the document must exist.    
    let getExcelDoc (absolutePath:string) : DocMonad<ExcelDoc, 'userRes> = 
        getDocument [".xls"; ".xlsx"; ".xlsm"] absolutePath

    /// Must have .xls or .xlsx or .xlsm extension. 
    /// Condition - the document must exist.
    let getWorkingExcelDoc (relativeName:string) : DocMonad<ExcelDoc, 'userRes> = 
        getWorkingDocument [".xls"; ".xlsx"; ".xlsm"] relativeName


    /// Must have .xls or .xlsx or .xlsm extension.
    /// Condition - the document must exist.
    let getSourceExcelDoc (relativeName:string) : DocMonad<ExcelDoc, 'userRes> = 
        getSourceDocument [".xls"; ".xlsx"; ".xlsm"] relativeName

    /// Must have .xls or .xlsx or .xlsm extension. 
    /// Condition - the document must exist.
    let getIncludeExcelDoc (relativeName:string) : DocMonad<ExcelDoc, 'userRes> = 
        getIncludeDocument [".xls"; ".xlsx"; ".xlsm"] relativeName


    // ************************************************************************
    // PowerPoint file (.ppt, .pptx)

    type PowerPointPhantom = class end

    type PowerPointDoc = Document<PowerPointPhantom>

    /// Must have .ppt or .pptx extension.
    /// Condition - the document must exist.    
    let getPowerPointDoc (absolutePath:string) : DocMonad<PowerPointDoc, 'userRes> = 
        getDocument [".ppt"; ".pptx"] absolutePath

    /// Must have .ppt or .pptx extension.
    /// Condition - the document must exist.
    let getWorkingPowerPointDoc (relativeName:string) : DocMonad<PowerPointDoc, 'userRes> = 
        getWorkingDocument [".ppt"; ".pptx"] relativeName 

    /// Must have .ppt or .pptx extension.
    /// Condition - the document must exist.
    let getSourcePowerPointDoc (relativeName:string) : DocMonad<PowerPointDoc, 'userRes> = 
        getSourceDocument [".ppt"; ".pptx"] relativeName 

    /// Must have .ppt or .pptx extension.
    /// Condition - the document must exist.
    let getIncludePowerPointDoc (relativeName:string) : DocMonad<PowerPointDoc, 'userRes> = 
        getIncludeDocument [".ppt"; ".pptx"] relativeName 


    // ************************************************************************
    // Text file (.txt)

    type TextPhantom = class end

    type TextDoc = Document<TextPhantom>

    /// Must have .txt extension.  
    let getTextDoc (absolutePath:string) : DocMonad<TextDoc, 'userRes> = 
        getDocument[".txt"] absolutePath


    /// Must have .txt extension.
    /// Condition - the document must exist.
    let getWorkingTextDoc (relativeName:string) : DocMonad<TextDoc, 'userRes> = 
        getWorkingDocument [".txt"] relativeName 

    /// Must have .txt extension.
    /// Condition - the document must exist.
    let getSourceTextDoc (relativeName:string) : DocMonad<TextDoc, 'userRes> = 
        getSourceDocument [".txt"] relativeName 

    /// Must have .txt extension.
    /// Condition - the document must exist.
    let getIncludeTextDoc (relativeName:string) : DocMonad<TextDoc, 'userRes> = 
        getIncludeDocument [".txt"] relativeName 
