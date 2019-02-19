// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause



module Proprietary

open FSharp.Interop.Excel
open ExcelProviderHelper

type ADBTable = 
    ExcelFile< FileName = @"G:\work\Projects\events2\ADB-all_sites_20181010.xlsx",
               SheetName = "Sheet1!",
               ForceString = true >

type ADBRow = ADBTable.Row

let readADBAll () : ADBRow list = 
    let helper = 
        { new IExcelProviderHelper<ADBTable, ADBRow>
          with member this.ReadTableRows table = table.Data 
               member this.IsBlankRow row = match row.GetValue(0) with null -> true | _ -> false }
    new ADBTable() |> excelReadRowsAsList helper



    