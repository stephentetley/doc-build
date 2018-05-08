#I @"..\packages\ExcelProvider.0.8.2\lib"
#r "ExcelProvider.dll"
open FSharp.ExcelProvider

#I @"..\packages\Newtonsoft.Json.10.0.3\lib\net45"
#r "Newtonsoft.Json"
open Newtonsoft.Json

// FAKE is local to the project file
#I @"..\packages\FAKE.5.0.0-beta005\tools"
#r @"..\packages\FAKE.5.0.0-beta005\tools\FakeLib.dll"

open System

#load @"DocMake\Base\Common.fs"
#load @"DocMake\Base\JsonUtils.fs"
#load @"DocMake\Base\GENHelper.fs"
open DocMake.Base.Common
open DocMake.Base.JsonUtils
open DocMake.Base.GENHelper

type AssetTable = 
    ExcelFile< @"G:\work\Projects\plcar\final_docs\KW-Batch01.xlsx",
               SheetName = "Sheet1",
               ForceString = false >

type AssetRow = AssetTable.Row


let getAssetRows () : AssetRow list = 
    let rowsDict : GetRowsDict<AssetTable, AssetRow> = 
        { GetRows     = fun imports -> imports.Data 
          NotNullProc = fun row -> match row.GetValue(0) with | null -> false | _ -> true }
    excelTableGetRows rowsDict (new AssetTable())

let makeDict (row:AssetRow) : FindReplaceDict = 
    Map.ofList [ "#SITENAME",   row.Site
               ; "#SAINUM" ,    row.SAI
               ; "#PLC",        row.PLC
               ]

let jsonConfig : FindsReplacesConfig<AssetRow> = 
    { DictionaryBuilder = makeDict
    ; GetFileName       = 
        fun (row:AssetRow) -> sprintf "%s_findreplace.json" (safeName row.Folder)

    ; OutputJsonFolder = @"G:\work\Projects\kw_plcar\final_docs\__Json" }

let batchConfig : BatchFileConfig = 
    { PathToFake = @"D:\coding\fsharp\DocMake\packages\FAKE.5.0.0-beta005\tools\FAKE.exe"
      PathToScript = @"D:\coding\fsharp\DocMake\src\PLCAR_Final_Build.fsx"
      BuildTarget = "Final"
      VarName = "assetname"
      OutputBatchFile = @"G:\work\Projects\kw_plcar\final_docs\fake-make.bat" }


// A file is generated foreach row
let main () : unit = 
    let assetList = getAssetRows () 
    // Generate find-replace json...
    assetList |> List.iter (generateFindsReplacesJson jsonConfig)
    // Generate batch file...
    assetList 
        |> List.map (fun (row:AssetRow) -> row.Folder) 
        |> generateBatchFile batchConfig 