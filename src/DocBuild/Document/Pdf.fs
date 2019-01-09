// Copyright (c) Stephen Tetley 2018,2019
// License: BSD 3 Clause


namespace DocBuild.Document

open DocBuild.Raw.Ghostscript.Ghostscript


// This should support extraction / rotation via Pdftk...

[<AutoOpen>]
module Pdf = 

    open DocBuild.Base
    open DocBuild.Base.Shell
    open DocBuild.Base.Monad
    open DocBuild.Base.Document
    open DocBuild.Base.Collective
    open DocBuild.Raw.Ghostscript
    open DocBuild.Raw.Pdftk
    open DocBuild.Raw.PdftkRotate
    
    

    type PdfFile = 
        val private PdfDoc : Document

        internal new (doc:Document) = 
            { PdfDoc = doc }

        new (filePath:string) = 
            { PdfDoc = new Document(filePath = filePath) }


        member internal x.Document 
            with get() : Document = x.PdfDoc

        member x.ActiveFile
            with get() : FilePath = x.PdfDoc.ActiveFile 

        member x.SaveAs(outputPath: string) : unit =  
            x.PdfDoc.SaveAs(outputPath)
            
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

    let pdfFile (path:string) : PdfFile = new PdfFile (filePath = path)

    
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



    type PdfColl = 
        val private Pdfs : Collective

        new (pdfs:PdfFile list) = 
            let docs = pdfs |> List.map (fun x -> x.Document)
            { Pdfs = new Collective(docs = docs) }

        member x.Documents 
            with get() : PdfFile list = 
                x.Pdfs.Documents |> List.map (fun d -> new PdfFile(doc=d))


    let pdfColl (pdfs:PdfFile list) : PdfColl = 
        new PdfColl(pdfs=pdfs)



    let ghostscriptConcat (inputfiles:PdfColl)
                            (quality:GsQuality)
                            (outputFile:string) : DocBuild<string> = 
            let inputs = inputfiles.Documents |> List.map (fun d -> d.ActiveFile)
            let cmd = makeGsConcatCommand quality.QualityArgs outputFile inputs
            execGhostscript cmd