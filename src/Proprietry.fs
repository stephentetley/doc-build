// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

// This file is functions and utilities to work with proprietry (clients) data.
// It is essentially "uninteresting".

module Proprietry

open FSharp.Data
open FSharp.ExcelProvider

open DocMake.Base.Common

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

let slashName (siteName:string) : string =
    siteName.Replace("_","/")

let underscoreName (siteName:string) : string = 
    siteName.Replace("/","_")
