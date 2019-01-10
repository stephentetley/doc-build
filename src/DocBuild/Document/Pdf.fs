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

    

    type RotationDirection = PdftkPrim.RotationDirection

    type RotationDirective = PdftkPrim.RotationDirective

    let rotSinglePage (pageNumber:int) (direction:RotationDirection) : RotationDirective = 
        { StartPage = pageNumber
          EndPage = pageNumber
          Direction = direction
        }
    
    let rotToEnd (startPage:int) (direction:RotationDirection) : RotationDirective = 
        { StartPage = startPage
          EndPage = -1
          Direction = direction
        }
    
    let rotRange (startPage:int) (endPage:int) (direction:RotationDirection) : RotationDirective = 
        { StartPage = startPage
          EndPage = endPage
          Direction = direction
        }


        //member x.RotateEmbed( options:ProcessOptions
        //                    , rotations: Rotation list)  : unit = 
        //    match pdfRotateEmbed options rotations x.PdfDoc.ActiveFile x.PdfDoc.ActiveFile with
        //    | ProcSuccess _ -> ()
        //    | ProcErrorCode i -> 
        //        failwithf "PdfDoc.RotateEmbed - error code %i" i
        //    | ProcErrorMessage msg -> 
        //        failwithf "PdfDoc.RotateEmbed - '%s'" msg
                
        //member x.RotateExtract( options:ProcessOptions
        //                      , rotations: Rotation list)  : unit = 
        //    match pdfRotateExtract options rotations x.PdfDoc.ActiveFile x.PdfDoc.ActiveFile with
        //    | ProcSuccess _ -> ()
        //    | ProcErrorCode i -> 
        //        failwithf "PdfDoc.RotateEmbed - error code %i" i
        //    | ProcErrorMessage msg -> 
        //        failwithf "PdfDoc.RotateEmbed - '%s'" msg




    // let pdfPageCount (pdf:PdfFile) : DocMonad<int> = 

    

    
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






    let ghostscriptConcat (inputfiles:PdfFile list)
                            (quality:GsQuality)
                            (outputFile:string) : DocMonad<string> = 
        let inputs = inputfiles |> List.map (fun d -> d.Path)
        let cmd = GhostscriptPrim.concatCommand quality.QualityArgs outputFile inputs
        execGhostscript cmd

    let pdfPageCount (inputfile:PdfFile) : DocMonad<int> = 
        docMonad { 
            let command = PdftkPrim.dumpDataCommand inputfile.Path
            let! stdout = execPdftk command
            let! ans = liftResult (PdftkPrim.regexSearchNumberOfPages stdout)
            return ans
        }
