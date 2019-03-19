// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Raw


// TODO
// Use pdftk for an alternative non-optimizing `pdf concat`.
// It appears Ghostscript concat can be too aggressive with 
// optimization.
// We should only optimize on a final pass.



[<RequireQualifiedAccess>]
module PdftkPrim = 

    open System.Text.RegularExpressions
    
    open SLFormat.CommandOptions

    open DocBuild.Base
    open DocBuild.Base.Shell

    

    /// Apparently we cannot send multiline commands to execProcess.
    let dumpDataCommand (inputFile: string) : CmdOpt list = 
        [ literal (doubleQuote inputFile) ; argument "dump_data" ]


    /// Seacrh for number of pages in a dump_data from Pdftk
    /// NumberOfPages: 3
    let regexSearchNumberOfPages (dumpData:string) : Result<int,ErrMsg> = 
        let patt = @"NumberOfPages: (\d+)"
        let result = Regex.Match(dumpData, patt)
        if result.Success then 
                result.Groups.Item(1).Value |> int |> Ok
        else 
            Error "regexSearchNumberOfPages 'NumberOfPages' not found"
        

    // ************************************************************************
    // Rotation

    type RotationDirection = 
        RotateNorth | RotateSouth | RotateEast | RotateWest
        member internal v.DirectionName
            with get () = 
                match v with
                | RotateNorth -> "north"
                | RotateSouth -> "south"
                | RotateEast -> "east"
                | RotateWest -> "west"
    
    type RotationDirective = 
        { StartPage: int
          EndPage: int
          Direction: RotationDirection }


    let rotationSpec (directive:RotationDirective): CmdOpt = 
        let direction = directive.Direction.DirectionName
        match directive.StartPage, directive.EndPage with
        | s,t when s < t -> 
            literal (sprintf "%i-%i%s" s t direction)
        | s,t when s = t ->
            literal (sprintf "%i%s" s direction)
        | s,t ->
            literal (sprintf "%i-end%s" s direction)

    let rotationSpecs (directives:RotationDirective list): CmdOpt list = 
        List.map rotationSpec directives

    /// <inputFile> cat <directive1> ... output <outputFile>
    let rotationCommand (inputFile: string) 
                        (directives:RotationDirective list)
                        (outputFile: string) : CmdOpt list = 
        let aa = [ literal (doubleQuote inputFile) ; argument "cat"] 
        let bb = rotationSpecs directives 
        let cc = [ argument "output"; literal (doubleQuote outputFile) ]
        aa @ bb @ cc

    /// To do - look at old pdftk rotate and redo the code 
    /// for rotating just islands in a document (keeping the water)
