// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Raw


// TODO - can use pdftk to get number of pages:
// > pdftk mydoc.pdf dump_data 
// Look for NumberOfPages in the output


[<RequireQualifiedAccess>]
module PdftkPrim = 

    open System.Text.RegularExpressions

    open DocBuild.Base
    open DocBuild.Base.Shell

    

    /// Apparently we cannot send multiline commands to execProcess.
    let dumpDataCommand (inputFile: string) : CommandArgs = 
        noArg (doubleQuote inputFile) ^^ noArg "dump_data"


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


    let rotationRange (directive:RotationDirective): CommandArgs = 
        let direction = directive.Direction.DirectionName
        match directive.StartPage, directive.EndPage with
        | s,t when s < t -> 
            noArg (sprintf "%i-%i%s" s t direction)
        | s,t when s = t ->
            noArg (sprintf "%i%s" s direction)
        | s,t ->
            noArg (sprintf "%i-end%s" s direction)