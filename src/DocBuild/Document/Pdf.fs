// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Document


// Supports rotation via Pdftk, Pdftk page extraction to add...

module Pdf = 
    
    open System

    open SLFormat.CommandOptions

    open DocBuild.Base
    open DocBuild.Base.DocMonad
    open DocBuild.Raw

    // ************************************************************************
    // Concatenation
    

    /// Concatenate a collection of Pdfs into a single Pdf with PDFtk.
    /// The result is output in the working directory.
    /// This produces large files than Ghostscipt with the option
    /// `/default`.
    let pdftkConcatPdfs (inputFiles:PdfCollection)
                        (outputAbsPath:string) : DocMonad<'userRes,PdfDoc> = 
        docMonad { 
            do! assertIsWorkingPath outputAbsPath
            let inputs = 
                inputFiles.Elements |> List.map (fun d -> d.AbsolutePath)
            let cmd = PdftkPrim.concatCommand inputs outputAbsPath
            let! _ = execPdftk cmd
            return! getWorkingPdfDoc outputAbsPath
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
                                  (inputFiles:PdfCollection) : DocMonad<'userRes,string> = 
        let inputs = 
            inputFiles.Elements |> List.map (fun d -> d.AbsolutePath)
        let cmd = GhostscriptPrim.concatCommand quality.QualityArgs outputAbsPath inputs
        execGhostscript cmd

    /// Concatenate a collection of Pdfs into a single Pdf
    /// with Ghostscript.
    /// The result is output in the working directory.
    let concatPdfs (quality:GsQuality)
                   (inputFiles:PdfCollection)
                   (outputAbsPath:string) : DocMonad<'userRes,PdfDoc> = 
        docMonad { 
            do! assertIsWorkingPath outputAbsPath
            let! _ = ghostscriptConcat quality outputAbsPath inputFiles
            return! getWorkingPdfDoc outputAbsPath
 
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
                           (outputAbsPath:string) 
                           (src:PdfDoc) : DocMonad<'userRes,PdfDoc> = 
        docMonad { 
            do! assertIsWorkingPath outputAbsPath
            let command = 
                PdftkPrim.rotationCommand src.AbsolutePath directives outputAbsPath
            let! _ = execPdftk command
            return! getWorkingPdfDoc outputAbsPath
        }

    /// Rezize for Word generating a new temp file
    let extractRotations (directives:RotationDirective list) 
                         (src:PdfDoc) : DocMonad<'userRes,PdfDoc> = 
        extractRotationsAs directives src.AbsolutePath src




    /// To do - look at old pdftk rotate and redo the code 
    /// for rotating just islands in a document (keeping the water)


    // ************************************************************************
    // Page count

    let countPages (inputfile:PdfDoc) : DocMonad<'userRes,int> = 
        docMonad { 
            let command = PdftkPrim.dumpDataCommand inputfile.AbsolutePath
            let! stdout = execPdftk command
            return! liftResult (PdftkPrim.regexSearchNumberOfPages stdout)
        }


    let sumPages (col:PdfCollection) : DocMonad<'userRes,int> = 
        mapM countPages col.Elements |>> List.sum