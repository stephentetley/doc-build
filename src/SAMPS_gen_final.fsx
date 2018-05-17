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

type SitesTable = 
    ExcelFile< @"G:\work\Projects\samps\final-docs\input\May2018_batch03\site_list.xlsx",
               SheetName = "Site_List",
               ForceString = false >

type SitesRow = SitesTable.Row

let sitesTableDict : GetRowsDict<SitesTable, SitesRow> = 
    { GetRows     = fun imports -> imports.Data 
      NotNullProc = fun row -> match row.GetValue(0) with | null -> false | _ -> true }

let getSitesRows () : SitesRow list = excelTableGetRows sitesTableDict (new SitesTable())



let makeDict (row:SitesRow) : FindReplaceDict = 
    Map.ofList [ "#SITENAME", row.Site
               ; "#SAINUM" , row.UID
               ]

let jsonConfig : FindsReplacesConfig<SitesRow> = 
    { DictionaryBuilder = makeDict
    ; GetFileName       = 
        fun (row:SitesRow) -> sprintf "%s_findreplace.json" (safeName row.Site)

    ; OutputJsonFolder = @"G:\work\Projects\samps\final-docs\__Json" }

let batchConfig : BatchFileConfig = 
    { PathToFake = @"D:\coding\fsharp\DocMake\packages\FAKE.5.0.0-beta005\tools\FAKE.exe"
      PathToScript = @"D:\coding\fsharp\DocMake\src\SAMPS_Final_Build_Alt.fsx"
      BuildTarget = "Final"
      OutputBatchFile = @"G:\work\Projects\samps\final-docs\fake-make.bat"
      VarName = "sitename" }


// A file is generated foreach row
let main () : unit = 
    let siteList = getSitesRows () 
    // Generate find-replace json...
    siteList |> List.iter (generateFindsReplacesJson jsonConfig)
    // Generate batch file...
    siteList 
        |> List.map (fun (row:SitesRow) -> row.Site) 
        |> generateBatchFile batchConfig 