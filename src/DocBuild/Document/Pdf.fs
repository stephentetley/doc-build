// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Document



// This should support extraction / rotation via Pdftk...

[<AutoOpen>]
module Pdf = 
    
    open System

    open DocBuild.Base
    open DocBuild.Base.Shell
    open DocBuild.Base.DocMonad
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
                                  (outputFile:string) 
                                  (inputFiles:PdfCollection) : DocMonad<'res,string> = 
        let inputs = inputFiles |> Collection.toList |> List.map (fun d -> d.Path)
        let cmd = GhostscriptPrim.concatCommand quality.QualityArgs outputFile inputs
        execGhostscript cmd


    let pdfConcat (quality:GsQuality)
                  (outputFile:string) 
                  (inputFiles:PdfCollection): DocMonad<'res,PdfFile> = 
        docMonad { 
            let! _ = ghostscriptConcat quality outputFile inputFiles
            let! pdf = getPdfFile outputFile
            return pdf
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



    let extractRotationsAs (directives:RotationDirective list)
                           (outputFile:string) 
                           (src:PdfFile) : DocMonad<'res,PdfFile> = 
        docMonad { 
            let command = 
                PdftkPrim.rotationCommand src.Path directives outputFile
            let! _ = execPdftk command
            let! pdf = getPdfFile outputFile
            return pdf
        }

    /// Rezize for Word generating a new temp file
    let extractRotations (directives:RotationDirective list) 
                         (src:PdfFile) : DocMonad<'res,PdfFile> = 
        extractRotationsAs directives src.NextTempName src




    /// To do - look at old pdftk rotate and redo the code 
    /// for rotating just islands in a document (keeping the water)


    // ************************************************************************
    // Page count

    let pdfPageCount (inputfile:PdfFile) : DocMonad<'res,int> = 
        docMonad { 
            let command = PdftkPrim.dumpDataCommand inputfile.Path
            let! stdout = execPdftk command
            let! ans = liftResult (PdftkPrim.regexSearchNumberOfPages stdout)
            return ans
        }
