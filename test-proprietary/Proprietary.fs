// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause



module Proprietary

open FSharp.Interop.Excel
open ExcelProviderHelper



type ADBTable = 
    ExcelFile< FileName = @"G:\work\ADB-exports\ADB-all_sites_20181010.xlsx",
               SheetName = "Sheet1!",
               ForceString = true >

type ADBRow = ADBTable.Row

let readADBAll () : ADBRow list = 
    let helper = 
        { new IExcelProviderHelper<ADBTable, ADBRow>
          with member this.ReadTableRows table = table.Data 
               member this.IsBlankRow row = match row.GetValue(0) with null -> true | _ -> false }
    new ADBTable() |> excelReadRowsAsList helper



type SaiMap = Map<string,string> 

let buildSaiMap () : SaiMap = 
    let rows = readADBAll () 
    List.fold (fun acc (row:ADBRow) -> 
                    Map.add row.InstCommonName row.InstReference acc) 
              Map.empty
              rows
    

let getSaiNumber (saiMap:SaiMap) (siteName:string) : string option = 
    match Map.tryFind siteName saiMap with
    | None -> printfn "Could not find SAI for '%s'" siteName ; None
    | Some ans -> Some ans

