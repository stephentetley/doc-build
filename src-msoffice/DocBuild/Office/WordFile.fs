// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Office


[<RequireQualifiedAccess>]
module WordFile = 

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
            let! ans = mf wordHandle.WordExe
            return ans
        }

    // ************************************************************************
    // Export

    let private wordExportQuality (quality:PrintQuality) : Word.WdExportOptimizeFor = 
        match quality with
        | PqScreen -> Word.WdExportOptimizeFor.wdExportOptimizeForOnScreen
        | PqPrint -> Word.WdExportOptimizeFor.wdExportOptimizeForPrint


    let exportPdfAs (quality:PrintQuality) 
                    (outputFile:string)
                    (src:WordFile) : DocMonad<#HasWordHandle,PdfFile> = 
        docMonad { 
            let pdfQuality = wordExportQuality quality
            let! (ans:unit) = 
                execWord <| fun app -> 
                    liftResult (wordExportAsPdf app src.Path pdfQuality outputFile)
            let! pdf = getPdfFile outputFile
            return pdf
        }

    /// Saves the file in the working directory.
    let exportPdf (quality:PrintQuality)  
                  (src:WordFile) : DocMonad<#HasWordHandle,PdfFile> = 
        docMonad { 
            let! local = Path.GetFileName(src.Path) |> changeToWorkingFile
            let outputFile = Path.ChangeExtension(local, "pdf")
            let! pdf = exportPdfAs quality outputFile src
            return pdf
        }



    // ************************************************************************
    // Find and replace

    let findReplaceAs (src:WordFile) (searches:SearchList) (outputFile:string) : DocMonad<#HasWordHandle,WordFile> = 
        docMonad { 
            let! ans = 
                execWord <| fun app -> 
                        liftResult (wordFindReplace app src.Path outputFile searches)
            let! docx = getWordFile outputFile
            return docx
        }



    let findReplace (src:WordFile) (searches:SearchList)  : DocMonad<#HasWordHandle,WordFile> = 
        findReplaceAs src searches src.NextTempName 
