module DocMake.Base.Common

open System.IO
open System.Text
open Fake.Core.Globbing.Operators


type FindReplaceDict = Map<string,string>


type DocMakePrintQuality = PqScreen | PqPrint


/// Use PdfWhatever for no optimization (can be faster and smaller in some cases...)
type PdfPrintSetting = PdfPrint | PdfScreen | PdfWhatever

let ghostscriptPrintSetting (quality:PdfPrintSetting) : string = 
    match quality with
    | PdfScreen -> @"-dPDFSETTINGS=/screen"
    | PdfPrint -> @"-dPDFSETTINGS=/preprint"
    | PdfWhatever -> ""

type DocMakePageOrientation = PoNorth | PoSouth | PoEast | PoWest

let pdftkPageOrientation (orientation:DocMakePageOrientation) : string = 
    match orientation with
    | PoNorth -> @"north"
    | PoSouth -> @"south"
    | PoEast -> @"east"
    | PoWest -> @"west"


let doubleQuote (s:string) : string = "\"" + s + "\""

let safeName (input:string) : string = 
    let bads = ['\\'; '/'; ':']
    List.fold (fun s c -> s.Replace(c,'_')) input bads

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
 
