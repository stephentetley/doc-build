// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Office


[<RequireQualifiedAccess>]
module WordDocument = 

    open System.IO

    // Open at .Interop rather than .Word then the Word API has to be qualified
    open Microsoft.Office.Interop


    open DocBuild.Base
    open DocBuild.Base.DocMonad
    open DocBuild.Base.DocMonadOperators
    open DocBuild.Office.Internal
    open Microsoft.Office.Interop.Word

    type WordHandle = 
        val mutable private WordApplication : Word.Application 
        val mutable private WordPaperSize : Word.WdPaperSize option
        new () = 
            { WordApplication = null; WordPaperSize = None }

        /// Opens a handle as needed.
        member x.WordExe : Word.Application = 
            match x.WordApplication with
            | null -> 
                let word1 = initWord ()
                x.WordApplication <- word1
                word1
            | app -> app


        interface ResourceFinalize with
            member x.RunFinalizer = 
                match x.WordApplication with
                | null -> () 
                | app -> finalizeWord app
    
        interface HasWordHandle with
            member x.WordAppHandle = x
            member x.PaperSizeForWord 
                with get () = x.WordPaperSize
                and set(v) = x.WordPaperSize <- v

    

    and HasWordHandle =
        abstract WordAppHandle : WordHandle
        abstract PaperSizeForWord : Word.WdPaperSize option with get, set

    let execWord (mf: Word.Application -> DocMonad<#HasWordHandle,'a>) : DocMonad<#HasWordHandle,'a> = 
        docMonad { 
            let! userRes = askUserResources ()
            let wordHandle = userRes.WordAppHandle
            return! mf wordHandle.WordExe
        }

    // ************************************************************************
    // Export

    let private wordExportQuality (quality:PrintQuality) : Word.WdExportOptimizeFor = 
        match quality with
        | Screen -> Word.WdExportOptimizeFor.wdExportOptimizeForOnScreen
        | Print -> Word.WdExportOptimizeFor.wdExportOptimizeForPrint


    let exportPdfAs (outputAbsPath:string)
                    (src:WordDoc) : DocMonad<#HasWordHandle,PdfDoc> = 
        docMonad { 
            do! assertIsWorkingPath outputAbsPath
            let! userRes = askUserResources ()
            let paperSize = userRes.PaperSizeForWord
            let! pdfQuality = 
                asks (fun env -> env.PrintOrScreen) |>> wordExportQuality
            let! (ans:unit) = 
                execWord <| fun app -> 
                    liftResult (wordExportAsPdf app paperSize pdfQuality src.LocalPath outputAbsPath)
            return! workingPdfDoc outputAbsPath
        }

    /// Saves the file in the top-level working directory.
    let exportPdf (src:WordDoc) : DocMonad<#HasWordHandle,PdfDoc> = 
        docMonad { 
            let! path1 = extendWorkingPath src.FileName
            let outputAbsPath = Path.ChangeExtension(path1, "pdf")
            return! exportPdfAs outputAbsPath src
        }



    // ************************************************************************
    // Find and replace

    let findReplaceAs (searches:SearchList) 
                      (outputAbsPath:string) 
                      (src:WordDoc) : DocMonad<#HasWordHandle,WordDoc> = 
        docMonad { 
            do! assertIsWorkingPath outputAbsPath
            let! ans = 
                execWord <| fun app -> 
                        liftResult (wordFindReplace app searches src.LocalPath outputAbsPath)
            return! workingWordDoc outputAbsPath
        }



    let findReplace (searches:SearchList) (src:WordDoc) : DocMonad<#HasWordHandle,WordDoc> = 
        findReplaceAs searches src.LocalPath src 
