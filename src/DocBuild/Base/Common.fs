// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


namespace DocBuild.Base

[<AutoOpen>]
module Common = 

    open System.IO


    type PdftkOptions = 
        { WorkingDirectory: string 
          PdftkExe: string 
        }

    type GsPdfSettings = 
        | GsPdfScreen 
        | GsPdfEbook
        | GsPdfPrinter
        | GsPdfPrepress
        | GsPdfDefault
        | GsPdfNone
        member v.PrintSetting 
            with get() = 
                match v with
                | GsPdfScreen ->  @"/screen"
                | GsPdfEbook -> @"/ebook"
                | GsPdfPrinter -> @"/printer"
                | GsPdfPrepress -> @"/prepress"
                | GsPdfDefault -> @"/default"
                | GsPdfNone -> ""

    type GhostscriptOptions = 
        { WorkingDirectory: string 
          GhostscriptExe: string 
          PrintQuality: GsPdfSettings
        }


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

    // ************************************************************************
    // Find and replace


    type SearchList = (string * string) list


    // ************************************************************************
    // File name helpers


    let safeName (input:string) : string = 
        let parens = ['('; ')'; '['; ']'; '{'; '}']
        let bads = ['\\'; '/'; ':'; '?'; '*'] 
        let white = ['\n'; '\t']
        let ans1 = List.fold (fun (s:string) (c:char) -> s.Replace(c.ToString(), "")) input parens
        let ans2 = List.fold (fun (s:string) (c:char) -> s.Replace(c,'_')) ans1 bads
        let ans3 = List.fold (fun (s:string) (c:char) -> s.Replace(c,'_')) ans2 white
        ans3.Trim() 


    /// Suffix a file name _before_ the extension.
    ///
    /// e.g suffixFileName "TEMP"  "sunset.jpg" ==> "sunset.TEMP.jpg"
    let suffixFileName (suffix:string)  (filePath:string) : string = 
        let root = System.IO.Path.GetDirectoryName filePath
        let justfile = System.IO.Path.GetFileNameWithoutExtension filePath
        let ext  = System.IO.Path.GetExtension filePath
        let newfile = sprintf "%s.%s%s" justfile suffix ext
        Path.Combine(root, newfile)

    // ************************************************************************
    // RunProcess = 

    /// Return `Choice2Of2 0` indicates Success
    let executeProcess (workingDirectory:string) (toolPath:string) (command:string) : Choice<string,int> = 
        try
            let procInfo = new System.Diagnostics.ProcessStartInfo ()
            procInfo.WorkingDirectory <- workingDirectory
            procInfo.FileName <- toolPath
            procInfo.Arguments <- command
            procInfo.CreateNoWindow <- true
            let proc = new System.Diagnostics.Process()
            proc.StartInfo <- procInfo
            proc.Start() |> ignore
            proc.WaitForExit () 
            Choice2Of2 <| proc.ExitCode
        with
        | ex -> Choice1Of2 (sprintf "executeProcess: \n%s" ex.Message)