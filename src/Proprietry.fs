// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

// This file contains functions and utilities to work with proprietry data 
// (usually spreadsheets).
// It is essentially "uninteresting".

module Proprietry

open FSharp.Data
open FSharp.ExcelProvider

open DocMake.Base.Common
open DocMake.Base.FakeLike
open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis



// *************************************
// SAI numbers


// SAI numbers are the uids for a proprietry data set we use.
// There is nothing interesting about them, they just form a dictionary - name-to-uid.

[<Literal>]
let SaiFilePath = __SOURCE_DIRECTORY__ + @"\..\data\SaiNumbers.xlsx"

type SaiTable = 
    ExcelFile< SaiFilePath,
                SheetName = "SAI_Data",
                ForceString = true >

type SaiRow = SaiTable.Row

let saiTableMethods : ExcelProviderHelperDict<SaiTable, SaiRow> = 
    { GetRows     = fun imports -> imports.Data 
      NotNullProc = fun row -> match row.GetValue(0) with null -> false | _ -> true }

let getSaiRows () : seq<SaiRow> = 
    excelTableGetRows saiTableMethods (new SaiTable())

type SaiLookups = Map<string, string>

let getSaiLookups () : SaiLookups = 
    let insert1 (ac:SaiLookups) (saiRow:SaiRow) = 
        Map.add saiRow.InstCommonName saiRow.InstReference ac

    getSaiRows () 
        |> Seq.fold insert1 Map.empty

let getSaiNumber (siteName:string) (saiLookups:SaiLookups) : option<string> = 
    Map.tryFind siteName saiLookups


// *************************************
// Name helpers

let noParens (input:string) : string = 
    let parens = ['('; ')'; '['; ']'; '{'; '}']
    let ans1 = List.fold (fun (s:string) (c:char) -> s.Replace(c.ToString(), "")) input parens
    ans1.Trim() 

/// Does not replace parens, squares, braces...
let slashName (siteName:string) : string =
    siteName.Replace("_","/")


/// Does not replace parens, squares, braces...
let underscoreName (siteName:string) : string = 
    siteName.Replace("/","_")



// *************************************
// Upload spreadsheet

// [<Literal>]
// let uploadSchema = @"Title(string), Sheet/Volume(string), Revision(string), Reference(string), Category(string), Project Name(string), File Format(string), File Date(string), Contractor(string)"
// let uploadHeaders = @"Title,Sheet/Volume,Revision,Reference,Category,Project Name,File Format,File Date,Contractor"

[<Literal>]
let UploadSchema = @"Asset Name(string), Asset Reference(string), Project Name(string), Project Code(string), Title(string), Category(string), Reference Number(string), Revision(string), Document Name(string), Document Date(string), Sheet/Volume(string)"
let uploadHeaders = @"Asset Name,Asset Reference,Project Name,Project Code,Title,Category,Reference Number,Revision,Document Name,Document Date,Sheet/Volume"



type UploadTable = 
    CsvProvider< Schema = UploadSchema,
                 HasHeaders = false >

type UploadRow = UploadTable.Row

let private outputUploadsCsv (source:seq<UploadRow>) (outputPath:string) : unit = 
    let table = new UploadTable(source)
    use sw = new System.IO.StreamWriter(outputPath)
    sw.WriteLine uploadHeaders
    table.Save(writer = sw, separator = ',', quote = '"' )


/// DocumentDate is usually: System.DateTime.Now.ToString("dd/MM/yyyy")
type UploadConstants = 
    { ProjectName: string           /// e.g. "RTU Asset Replacement"
      ProjectCode: string           /// Project code, e.g. "S3945"
      Category: string              /// e.g. "O & M Manuals"
      ReferenceNumber: string       /// Project code, e.g. "S3945"
      Revision:string               /// Usually "1"
      DocumentDate: string          /// Date in "dd/MM/yyyy" format
      SheetVolume: string }         /// Usually blank
 
let standardDocumentDate () : string  = System.DateTime.Now.ToString("dd/MM/yyyy") 

type UploadConfig = 
    { MakeTitle: string -> string
      MakeDocName: string -> string
      ConstantParams: UploadConstants }



let private uploadRow (siteName:string) (saiNum:option<string>) 
                (config:UploadConfig) : UploadRow = 
    match saiNum with
    | None -> 
        UploadTable.Row(assetName = siteName,
                        assetReference = "BAD",
                        projectName = "",
                        projectCode = "",
                        title = "",
                        category = "",
                        referenceNumber = "", 
                        revision = "",
                        documentName = "",
                        documentDate = "",
                        sheetVolume = "" )
    | Some saiNumber ->         
        UploadTable.Row(assetName = siteName,
                        assetReference = saiNumber,
                        projectName = config.ConstantParams.ProjectName,
                        projectCode = config.ConstantParams.ProjectCode,
                        title = config.MakeTitle siteName,
                        category = config.ConstantParams.Category,
                        referenceNumber = config.ConstantParams.ProjectCode, 
                        revision = config.ConstantParams.Revision,
                        documentName = config.MakeDocName siteName,
                        documentDate = config.ConstantParams.DocumentDate,
                        sheetVolume = config.ConstantParams.SheetVolume )

let makeUploadForm (siteNames:string list) 
                    (config:UploadConfig) : BuildMonad<'res,unit> = 
    buildMonad { 
        let saiDict = getSaiLookups ()
        let rows = 
            List.map (fun (name:string) -> 
                        uploadRow name (getSaiNumber name saiDict) config) siteNames
        let! cwd = askWorkingDirectory () 
        let outPath = cwd </> "__EDMS_Upload.csv"
        do! executeIO (fun () -> outputUploadsCsv rows outPath)
    }

                    
                        

