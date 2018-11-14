// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

module DocMake.Base.Common

open System.IO


type FindReplaceDict = Map<string,string>
type SearchList = (string*string) list



type PrintQuality = PqScreen | PqPrint


/// Use PdfWhatever for no optimization (can be faster and smaller in some cases...)
type PdfPrintQuality = PdfPrint | PdfScreen | PdfWhatever


// TODO options should be "/screen", "/ebook", "/printer", "/prepress", "/default"
let ghostscriptPrintSetting (quality:PdfPrintQuality) : string = 
    match quality with
    | PdfScreen -> @"/screen"
    | PdfPrint -> @"/prepress"      // WARNING To check
    | PdfWhatever -> ""

type PageOrientation = 
    PoNorth | PoSouth | PoEast | PoWest
        member v.PdftkOrientation = 
            match v with
            | PoNorth -> "north"
            | PoSouth -> "south"
            | PoEast -> "east"
            | PoWest -> "west"



// TODO - should escape, but that needs a new signature
let doubleQuote (s:string) : string = "\"" + s + "\""

let safeName (input:string) : string = 
    let parens = ['('; ')'; '['; ']'; '{'; '}']
    let bads = ['\\'; '/'; ':'; '?'] 
    let white = ['\n'; '\t']
    let ans1 = List.fold (fun (s:string) (c:char) -> s.Replace(c.ToString(), "")) input parens
    let ans2 = List.fold (fun (s:string) (c:char) -> s.Replace(c,'_')) ans1 bads
    let ans3 = List.fold (fun (s:string) (c:char) -> s.Replace(c,'_')) ans2 white
    ans3.Trim() 


let zeroPad (width:int) (value:int) = 
    let ss = value.ToString ()
    let diff = width - ss.Length
    String.replicate diff "0" + ss

let maybeCreateDirectory (dirpath:string) : unit = 
    if not <| Directory.Exists(dirpath) then 
        ignore <| Directory.CreateDirectory(dirpath)
    else ()

let unique (xs:seq<'a>) : 'a = 
    let next zs = match zs with
                    | [] -> failwith "unique - no matches."
                    | [z] -> z
                    | _ -> failwithf "unique - %i matches" zs.Length

    Seq.toList xs |> next

let pathChangeExtension (path:string) (extension:string) : string = 
    Path.ChangeExtension(path,extension)

let pathChangeDirectory (path:string) (outputdir:string) : string = 
    let justfile = Path.GetFileName path
    Path.Combine(outputdir,justfile)



 
