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
    open DocBuild.Office.Internal
    open Microsoft.Office.Interop.Word

    type WordHandle = 
        val mutable private WordApplication : Word.Application 
        val mutable private WordPaperSize : Word.WdPaperSize option
        new () = 
            { WordApplication = null
            ; WordPaperSize = Some WdPaperSize.wdPaperA4 }

        /// Opens a handle as needed.
        member x.WordExe : Word.Application = 
            match x.WordApplication with
            | null -> 
                let word1 = initWord ()
                x.WordApplication <- word1
                word1
            | app -> app


        interface IResourceFinalize with
            member x.RunFinalizer = 
                match x.WordApplication with
                | null -> () 
                | app -> finalizeWord app
    
        interface IWordHandle with
            member x.WordAppHandle = x
            member x.PaperSizeForWord 
                with get () = x.WordPaperSize
                and set(v) = x.WordPaperSize <- v

    

    and IWordHandle =
        abstract WordAppHandle : WordHandle
        abstract PaperSizeForWord : Word.WdPaperSize option with get, set

    let execWord (mf: Word.Application -> DocMonad<'a, #IWordHandle>) : DocMonad<'a, #IWordHandle> = 
        docMonad { 
            let! userRes = getUserResources ()
            let wordHandle = userRes.WordAppHandle
            return! mf wordHandle.WordExe
        }

    // ************************************************************************
    // Export

    let private wordExportQuality (quality:PrintQuality) : Word.WdExportOptimizeFor = 
        match quality with
        | Screen -> Word.WdExportOptimizeFor.wdExportOptimizeForOnScreen
        | Print -> Word.WdExportOptimizeFor.wdExportOptimizeForPrint


    /// TODO - because we only want output in the working directory
    /// it would be better to supply just a file name...
    let exportPdfAs (outputRelName:string)
                    (src:WordDoc) : DocMonad<PdfDoc, #IWordHandle> = 
        docMonad { 
            let! outputAbsPath = extendWorkingPath outputRelName
            let! userRes = getUserResources ()
            let paperSize = userRes.PaperSizeForWord
            let! pdfQuality = 
                asks (fun env -> env.PrintOrScreen) |>> wordExportQuality
            let! (ans:unit) = 
                execWord <| fun app -> 
                    liftOperationResult "exportAsPdf" (fun _ -> wordExportAsPdf app paperSize pdfQuality src.AbsolutePath outputAbsPath)
            return! getPdfDoc outputAbsPath
        }

    /// Saves the file in the top-level working directory.
    let exportPdf (src:WordDoc) : DocMonad<PdfDoc, #IWordHandle> = 
        docMonad { 
            let fileName = Path.ChangeExtension(src.FileName, "pdf")
            return! exportPdfAs fileName src
        }



    // ************************************************************************
    // Find and replace

    let findReplaceAs (searches:SearchList) 
                      (outputRelName:string) 
                      (src:WordDoc) : DocMonad<WordDoc, #IWordHandle> = 
        docMonad { 
            let! outputAbsPath = extendWorkingPath outputRelName
            let! ans = 
                execWord <| fun app -> 
                        liftOperationResult "findReplaceAs" (fun _ -> wordFindReplace app searches src.AbsolutePath outputAbsPath)
            return! getWordDoc outputAbsPath
        }



    let findReplace (searches:SearchList) (src:WordDoc) : DocMonad<WordDoc, #IWordHandle> = 
        findReplaceAs searches src.FileName src 
