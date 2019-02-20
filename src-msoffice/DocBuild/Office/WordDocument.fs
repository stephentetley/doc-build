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

    type WordHandle = 
        val mutable private WordApplication : Word.Application 

        new () = 
            { WordApplication = null }

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

    

    and HasWordHandle =
        abstract WordAppHandle : WordHandle

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
        | PqScreen -> Word.WdExportOptimizeFor.wdExportOptimizeForOnScreen
        | PqPrint -> Word.WdExportOptimizeFor.wdExportOptimizeForPrint


    let exportPdfAs (quality:PrintQuality) 
                    (outputAbsPath:string)
                    (src:WordDoc) : DocMonad<#HasWordHandle,PdfDoc> = 
        docMonad { 
            do! assertIsWorkingPath outputAbsPath
            let pdfQuality = wordExportQuality quality
            let! (ans:unit) = 
                execWord <| fun app -> 
                    liftResult (wordExportAsPdf app pdfQuality src.LocalPath outputAbsPath)
            return! workingPdfDoc outputAbsPath
        }

    /// Saves the file in the top-level working directory.
    let exportPdf (quality:PrintQuality)  
                  (src:WordDoc) : DocMonad<#HasWordHandle,PdfDoc> = 
        docMonad { 
            let! path1 = extendWorkingPath src.FileName
            let outputAbsPath = Path.ChangeExtension(path1, "pdf")
            return! exportPdfAs quality outputAbsPath src
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
