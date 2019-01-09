// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Document



// This should support extraction / rotation via Pdftk...

[<AutoOpen>]
module Pdf = 
    
    open System

    open DocBuild.Base
    open DocBuild.Base.Shell
    open DocBuild.Base.Monad
    open DocBuild.Base.Document
    open DocBuild.Base.Collective
    open DocBuild.Raw.Ghostscript
    open DocBuild.Raw.Pdftk
    open DocBuild.Raw.PdftkRotate
    
    
    [<Struct>]
    type PdfFile = 
        | PdfFile of Document

        member x.Path 
            with get () : FilePath =
                match x with | PdfFile(p) -> p.Path

        /// ActiveFile is a mutable working copy of the original file.
        /// The original file is untouched.
        member x.NextTempName
            with get() : FilePath = 
                match x with | PdfFile(p) -> p.NextTempName


    
            
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

    let pdfFile (path:string) : DocBuild<PdfFile> = 
        if System.IO.File.Exists(path) then 
            let extension : string = System.IO.Path.GetExtension(path)
            if String.Equals(extension, ".pdf", StringComparison.CurrentCultureIgnoreCase) then 
                breturn <| PdfFile(Document(path))
            else throwError <| sprintf "Not a pdf file: '%s'" path
        else throwError <| sprintf "Could not find file: '%s'" path  

    

    
    type GsQuality = 
        | GsScreen 
        | GsEbook
        | GsPrinter
        | GsPrepress
        | GsDefault
        | GsNone
        member v.QualityArgs
            with get() : CommandArgs = 
                match v with
                | GsScreen ->  reqArg "-dPDFSETTINGS" @"/screen"
                | GsEbook -> reqArg "-dPDFSETTINGS" @"/ebook"
                | GsPrinter -> reqArg "-dPDFSETTINGS" @"/printer"
                | GsPrepress -> reqArg "-dPDFSETTINGS" @"/prepress"
                | GsDefault -> reqArg "-dPDFSETTINGS" @"/default"
                | GsNone -> emptyArgs



    //type PdfColl = 
    //    val private Pdfs : Collective

    //    new (pdfs:PdfFile list) = 
    //        { Pdfs = new Collective(docs = pdfs) }

    //    member x.Documents 
    //        with get() : PdfFile list = 
    //            x.Pdfs.Documents |> List.map (fun d -> PdfFile(doc=d))


    //let pdfColl (pdfs:PdfFile list) : PdfColl = 
    //    new PdfColl(pdfs=pdfs)



    let ghostscriptConcat (inputfiles:PdfFile list)
                            (quality:GsQuality)
                            (outputFile:string) : DocBuild<string> = 
            let inputs = inputfiles |> List.map (fun d -> d.Path)
            let cmd = makeGsConcatCommand quality.QualityArgs outputFile inputs
            execGhostscript cmd