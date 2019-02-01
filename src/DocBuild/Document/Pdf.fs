﻿// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Document


// Supports rotation via Pdftk, extraction to add...

[<AutoOpen>]
module Pdf = 
    
    open System

    open DocBuild.Base
    open DocBuild.Base.Shell
    open DocBuild.Base.DocMonad
    open DocBuild.Base.DocMonadOperators
    open DocBuild.Raw

    // ************************************************************************
    // Concatenation
    
    type GsQuality = 
        | GsScreen 
        | GsEbook
        | GsPrinter
        | GsPrepress
        | GsDefault
        | GsNone
        member internal v.QualityArgs
            with get() : CommandArgs = 
                match v with
                | GsScreen ->  reqArg "-dPDFSETTINGS" @"/screen"
                | GsEbook -> reqArg "-dPDFSETTINGS" @"/ebook"
                | GsPrinter -> reqArg "-dPDFSETTINGS" @"/printer"
                | GsPrepress -> reqArg "-dPDFSETTINGS" @"/prepress"
                | GsDefault -> reqArg "-dPDFSETTINGS" @"/default"
                | GsNone -> emptyArgs






    let private ghostscriptConcat (quality:GsQuality)
                                  (outputAbsPath:string) 
                                  (inputFiles:PdfCollection) : DocMonad<'res,string> = 
        let inputs = 
            inputFiles |> Collection.toList |> List.map (fun d -> d.LocalPath)

        let cmd = GhostscriptPrim.concatCommand quality.QualityArgs outputAbsPath inputs
        execGhostscript cmd

    /// Concatenate a collection of Pdfs into a single Pdf.
    /// The result is output in the working directory.
    let pdfConcat (quality:GsQuality)
                  (outputAbsPath:string) 
                  (inputFiles:PdfCollection): DocMonad<'res,PdfFile> = 
        docMonad { 
            do! assertIsWorkingPath outputAbsPath
            let! _ = ghostscriptConcat quality outputAbsPath inputFiles
            return! workingPdfFile outputAbsPath
 
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
                           (src:PdfFile) : DocMonad<'res,PdfFile> = 
        docMonad { 
            do! assertIsWorkingPath outputAbsPath
            let command = 
                PdftkPrim.rotationCommand src.LocalPath directives outputAbsPath
            let! _ = execPdftk command
            return! workingPdfFile outputAbsPath
        }

    /// Rezize for Word generating a new temp file
    let extractRotations (directives:RotationDirective list) 
                         (src:PdfFile) : DocMonad<'res,PdfFile> = 
        extractRotationsAs directives src.LocalPath src




    /// To do - look at old pdftk rotate and redo the code 
    /// for rotating just islands in a document (keeping the water)


    // ************************************************************************
    // Page count

    let pdfPageCount (inputfile:PdfFile) : DocMonad<'res,int> = 
        docMonad { 
            let command = PdftkPrim.dumpDataCommand inputfile.LocalPath
            let! stdout = execPdftk command
            return! liftResult (PdftkPrim.regexSearchNumberOfPages stdout)
        }
