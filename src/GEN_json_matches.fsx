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
#load @"DocMake\Base\Json.fs"
open DocMake.Base.Common
open DocMake.Base.Json

#load @"GENHelper.fs"
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



let makeDict (row:SitesRow) : Dict = 
    Map.ofList [ "#SITENAME", row.Site
               ; "#SAINUM" , row.Uid
               ]

let jsonConfig : FindsReplacesConfig<SitesRow> = 
    { DictionaryBuilder = makeDict
    ; GetFileName       = 
        fun (row:SitesRow) -> sprintf "%s_findreplace.json" (safeName row.Site)

    ; OutputJsonFolder = @"G:\work\Projects\samps\final-docs\__Json" }


// A file is generated foreach row
let main () : unit = 
    getSitesRows () |> List.iter (generateFindsReplacesJson jsonConfig)