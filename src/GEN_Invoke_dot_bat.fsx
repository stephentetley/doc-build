#I @"..\packages\ExcelProvider.0.8.2\lib"
#r "ExcelProvider.dll"
open FSharp.ExcelProvider

open System

#load "GENHelper.fs"
open GENHelper

type SitesTable = 
    ExcelFile< @"G:\work\Projects\samps\sitelist-for-gen-jan2018.xlsx",
               SheetName = "Sheet1",
               ForceString = false >

type SitesRow = SitesTable.Row


let sitesTableDict : GetRowsDict<SitesTable, SitesRow> = 
    { GetRows     = fun imports -> imports.Data 
      NotNullProc = fun row -> match row.GetValue(0) with | null -> false | _ -> true }

let getSitesRows () : SitesRow list = excelTableGetRows sitesTableDict (new SitesTable())


let batchConfig : BatchFileConfig = 
    { PathToFake = @"D:\coding\fsharp\DocMake\packages\FAKE.5.0.0-beta005\tools\FAKE.exe"
      PathToScript = @"D:\coding\fsharp\DocMake\src\SAMPS_Final_Build.fsx"
      OutputBatchFile = @"G:\work\Projects\samps\fake-make.bat" }


let main () : unit = 
    generateBatchFile batchConfig 
        << List.map (fun (row:SitesRow) -> row.Site) 
        <| getSitesRows ()

