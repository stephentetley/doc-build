// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

module DocMake.Base.Common

open System.IO
open System.Text
open Fake.Core.Globbing.Operators


type FindReplaceDict = Map<string,string>
type SearchList = (string*string) list



type DocMakePrintQuality = PqScreen | PqPrint


/// Use PdfWhatever for no optimization (can be faster and smaller in some cases...)
type PdfPrintSetting = PdfPrint | PdfScreen | PdfWhatever


// TODO options should be "/screen", "/ebook", "/printer", "/prepress", "/default"
let ghostscriptPrintSetting (quality:PdfPrintSetting) : string = 
    match quality with
    | PdfScreen -> @"/screen"
    | PdfPrint -> @"/prepress"      // WARNING To check
    | PdfWhatever -> ""

type DocMakePageOrientation = PoNorth | PoSouth | PoEast | PoWest

let pdftkPageOrientation (orientation:DocMakePageOrientation) : string = 
    match orientation with
    | PoNorth -> @"north"
    | PoSouth -> @"south"
    | PoEast -> @"east"
    | PoWest -> @"west"

// TODO - should escape, but that needs a new signature
let doubleQuote (s:string) : string = "\"" + s + "\""

let safeName (input:string) : string = 
    let leftOf (find:char) (inp:string) : string = 
        match inp.Split(find) |> Array.toList with
        | (x ::_ ) -> x
        | _ -> inp
    let bads = ['\\'; '/'; ':'; '?']
    let ans1 = List.fold (fun (s:string) (c:char) -> s.Replace(c,'_')) input bads
    ans1.Trim() |> leftOf '[' |> leftOf '(' |> leftOf '{'


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


// The Excel Type Provider seems to read a trailing null row.
// This dictionary and procedure provide a skeleton to get round this.

type ExcelProviderHelperDict<'table, 'row> = 
    { GetRows : 'table -> seq<'row>
      NotNullProc : 'row -> bool }

let excelTableGetRows (dict:ExcelProviderHelperDict<'table,'row>) (table:'table) : 'row list = 
    let allrows = dict.GetRows table
    allrows |> Seq.filter dict.NotNullProc |> Seq.toList

 
