// Copyright (c) Stephen Tetley 2018, 2019
// License: BSD 3 Clause


namespace DocBuild.Raw.PdftkRotate

// TODO - can use pdftk to get number of pages:
// > pdftk mydoc.pdf dump_data 
// Look for NumberOfPages in the output


[<AutoOpen>]
module PdftkRotate = 

    open DocBuild.Base
    open DocBuild.Raw.Pdftk

    // ************************************************************************
    // Rotation

    type PageOrientation = 
        PoNorth | PoSouth | PoEast | PoWest
        member v.PdftkOrientation 
            with get () = 
                match v with
                | PoNorth -> "north"
                | PoSouth -> "south"
                | PoEast -> "east"
                | PoWest -> "west"

    type internal EndOfRange = 
        | EndOfDoc
        | EndPageNumber of int

    type Rotation = 
        internal 
            { StartPage: int
              EndPage: EndOfRange
              Orientation: PageOrientation }

    /// This is part of the API (should it need instantiating?)
    let rotationRange (startPage:int) (endPage:int) (orientation:PageOrientation) : Rotation = 
        { StartPage = startPage
          EndPage =  EndPageNumber endPage
          Orientation = orientation }

    let rotationSinglePage (pageNum:int) (orientation:PageOrientation) : Rotation = 
        rotationRange pageNum pageNum orientation

    /// This is part of the API (should it need instantiating?)
    let rotationToEnd (startPage:int) (orientation:PageOrientation) : Rotation = 
        { StartPage = startPage
          EndPage =  EndOfDoc
          Orientation = orientation }



    // Note - the rotate "API" of Pdftk appears cryptic:
    // It can be used to extract rotated pages or "embed" rotate pages within the 
    // rest of the document.
    // 
    // This is command embeds two rotated regions:
    // pdftk slides.pdf cat 1 2-4east 5 6-8east 9-end output slides-rot2.pdf
    // 
    // This command extracts two rotated regions:
    // pdftk slides.pdf cat 2-4east 6-8east output slides-rot3.pdf




    let private rotationSpec1 (rot:Rotation) : string = 
        let last = 
            match rot.EndPage with
            | EndOfDoc -> "end"
            | EndPageNumber i -> i.ToString()
        sprintf "%i-%s%s" rot.StartPage last rot.Orientation.PdftkOrientation

    let private regionSpec1 (startPage:int) (endPage:int) : string = 
        if endPage > startPage then 
            sprintf "%i-%i" startPage endPage
        else
            startPage.ToString()

    let private regionToEnd (startPage:int) : string = 
        sprintf "%i-end" startPage



    let private rotSpecForExtract (rotations: Rotation list) : string = 
        let rec work (ac:string list) rs = 
            match rs with
            | [] -> String.concat " " <| List.rev ac
            | rot1 :: rest -> work (rotationSpec1 rot1 :: ac) rest
        work [] rotations
            


    let private rotSpecForEmbed (rotations: Rotation list) : string = 
        let interRegion (prevStart:int) (nextStart:int) : option<string> = 
            let endOfRegion = nextStart - 1
            if endOfRegion >= prevStart then 
                Some <| regionSpec1 prevStart endOfRegion
            else None

        let rec work (page:int) (ac:string list) (rs:Rotation list) = 
            match rs with
            | [] -> 
                let final = regionToEnd page 
                in String.concat " " <| List.rev (final::ac)
            | rot1 :: rest -> 
                match rot1.EndPage with
                | EndOfDoc -> 
                    let final = rotationSpec1 rot1
                    let ac1 = 
                        match interRegion page rot1.StartPage with
                        | Some inter -> (final::inter::ac)
                        | None -> (final::ac)
                    String.concat " " ac1
                | EndPageNumber i ->
                    let next = rotationSpec1 rot1
                    let ac1 = 
                        match interRegion page rot1.StartPage with
                        | Some inter -> (next::inter::ac)
                        | None -> (next::ac)
                    work (i+1) ac1 rest
        work 1 [] rotations




    let private makeExtractCmd (inputFile:string) (outputFile:string) (rotations: Rotation list)  : string = 
        let rotateSpec = rotSpecForExtract rotations
        sprintf "\"%s\" cat %s output \"%s\"" inputFile rotateSpec outputFile 

    let private makeEmbedCmd (inputFile:string) (outputFile:string) (rotations: Rotation list)  : string = 
        let rotateSpec = rotSpecForEmbed rotations
        sprintf "\"%s\" cat %s output \"%s\"" inputFile rotateSpec outputFile 


    let pdfRotateExtract (options:PdftkOptions) 
                         (rotations: Rotation list) 
                         (inputFile:string) 
                         (outputFile:string) : Choice<string,int> =
        runPdftk options <| makeExtractCmd inputFile outputFile rotations



    let pdfRotateEmbed (options:PdftkOptions) 
                        (rotations: Rotation list) 
                        (inputFile:string) 
                        (outputFile:string) : Choice<string,int> =
        runPdftk options <|  makeEmbedCmd inputFile outputFile rotations







