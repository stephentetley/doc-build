// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Document


// Supports rotation via Pdftk, Pdftk page extraction to add...

module Pdf = 
    
    open System

    open SLFormat.CommandOptions

    open DocBuild.Base
    open DocBuild.Base.Internal
   

    // ************************************************************************
    // Concatenation
    

    /// Concatenate a collection of Pdfs into a single Pdf with PDFtk.
    /// The result is output in the working directory.
    /// This produces larger (but nicer) files than Ghostscipt with the option
    /// `/default`.
    let pdftkConcatPdfs (outputRelName:string)
                        (inputFiles:PdfCollection) : DocMonad<PdfDoc, 'userRes> = 
        docMonad { 
            let! outputAbsPath = extendWorkingPath outputRelName
            let inputs = inputFiles.DocumentPaths
            let cmd = PdftkPrim.concatCommand inputs outputAbsPath
            let! _ = execPdftk cmd
            return! getPdfDoc outputAbsPath
        }



    type GsQuality = 
        | GsScreen 
        | GsEbook
        | GsPrinter
        | GsPrepress
        | GsDefault
        | GsNone
        member internal v.QualityArgs
            with get() : CmdOpt = 
                match v with
                | GsScreen ->   argument "-dPDFSETTINGS"    &= @"/screen"
                | GsEbook ->    argument "-dPDFSETTINGS"    &= @"/ebook"
                | GsPrinter ->  argument "-dPDFSETTINGS"    &= @"/printer"
                | GsPrepress -> argument "-dPDFSETTINGS"    &= @"/prepress"
                | GsDefault ->  argument "-dPDFSETTINGS"    &= @"/default"
                | GsNone -> noArgument


    let private ghostscriptConcat (quality:GsQuality)
                                  (outputAbsPath:string) 
                                  (inputFiles:PdfCollection) : DocMonad<string, 'userRes> = 
        let inputs = inputFiles.DocumentPaths
        let cmd = GhostscriptPrim.concatCommand quality.QualityArgs outputAbsPath inputs
        execGhostscript cmd


    /// Concatenate a collection of Pdfs into a single Pdf
    /// with Ghostscript.
    /// The result is output in the working directory.
    /// Ghostscript is perhaps more tolerant of 'incorrect' PDF
    /// files than PDFtk.
    let concatPdfs (quality:GsQuality)
                   (outputRelName:string)
                   (inputFiles:PdfCollection) : DocMonad<PdfDoc, 'userRes> = 
        docMonad { 
            let! outputAbsPath = extendWorkingPath outputRelName
            let! _ = ghostscriptConcat quality outputAbsPath inputFiles
            return! getPdfDoc outputAbsPath
 
        }


    // ************************************************************************
    // Rotation

    type RotationDirection = PdftkPrim.RotationDirection

    type RotationDirective = PdftkPrim.RotationDirective

    let rotSinglePage (pageNumber:int) 
                      (direction:RotationDirection) : RotationDirective = 
        { StartPage = pageNumber
          EndPage = pageNumber
          Direction = direction
        }
    
    let rotToEnd (startPage:int) 
                 (direction:RotationDirection) : RotationDirective = 
        { StartPage = startPage
          EndPage = -1
          Direction = direction
        }
    
    let rotRange (startPage:int) 
                 (endPage:int) 
                 (direction:RotationDirection) : RotationDirective = 
        { StartPage = startPage
          EndPage = endPage
          Direction = direction
        }

    let private rotAll (direction:RotationDirection) : RotationDirective = 
        { StartPage = 1
          EndPage = -1
          Direction = direction
        }


    /// outputName is relatuive to Working directory.
    let extractRotationsAs (directives:RotationDirective list)
                           (outputRelName:string) 
                           (source : PdfDoc) : DocMonad<PdfDoc, 'userRes> = 
        docMonad { 
            let! outputAbsPath = extendWorkingPath outputRelName
            let! sourcePath = getDocumentPath source
            let command = 
                PdftkPrim.rotationCommand sourcePath directives outputAbsPath
            let! _ = execPdftk command
            return! getPdfDoc outputAbsPath
        }

    /// Resize for Word generating a new temp file
    let extractRotations (directives : RotationDirective list) 
                         (source : PdfDoc) : DocMonad<PdfDoc, 'userRes> = 
        docMonad { 
            let! sourceName = getDocumentFileName source
            return! extractRotationsAs directives sourceName source
        }




    /// To do - look at old pdftk rotate and redo the code 
    /// for rotating just islands in a document (keeping the water)


    // ************************************************************************
    // Page count

    let countPages (source:PdfDoc) : DocMonad<int, 'userRes> = 
        docMonad { 
            let! sourcePath = getDocumentPath source
            let command = PdftkPrim.dumpDataCommand sourcePath
            let! stdout = execPdftk command
            return! liftOperationResult "countPages" (fun _ -> PdftkPrim.regexSearchNumberOfPages stdout)
        }
        <|> mreturn 0


    let sumPages (col:PdfCollection) : DocMonad<int, 'userRes> = 
        mapM countPages col.Documents |>> List.sum