#I @"..\packages\ExcelProvider.0.8.2\lib"
#r "ExcelProvider.dll"
open FSharp.ExcelProvider


#I @"..\packages\Newtonsoft.Json.10.0.3\lib\net45"
#r "Newtonsoft.Json"
open Newtonsoft.Json

// FAKE is local to the project file
#I @"..\packages\FAKE.5.0.0-beta005\tools"
#r @"..\packages\FAKE.5.0.0-beta005\tools\FakeLib.dll"
open Fake
open Fake.Core


#load @"DocMake\Base\Common.fs"
#load @"DocMake\Base\JsonUtils.fs"
#load @"DocMake\Base\GENHelper.fs"
open DocMake.Base.GENHelper

// Only Generate Batch file 
// Find/Replace not needed


type SitesTable = 
    ExcelFile< @"G:\work\Projects\barriers\sites-temp.xlsx",
               SheetName = "Sites",
               ForceString = false >

type SitesRow = SitesTable.Row

let sitesTableDict : GetRowsDict<SitesTable, SitesRow> = 
    { GetRows     = fun imports -> imports.Data 
      NotNullProc = fun row -> match row.GetValue(0) with | null -> false | _ -> true }

let getSitesRows () : SitesRow list = excelTableGetRows sitesTableDict (new SitesTable())


let batchConfig : BatchFileConfig = 
    { PathToFake = @"D:\coding\fsharp\DocMake\packages\FAKE.5.0.0-beta005\tools\FAKE.exe"
      PathToScript = @"D:\coding\fsharp\DocMake\src\BARRIERS_Final_Build.fsx"
      BuildTarget = "Final"
      OutputBatchFile = @"G:\work\Projects\barriers\final-docs\fake-make.bat"
      VarName = "sitename" }

let main () : unit = 
    getSitesRows () 
        |> List.map (fun (row:SitesRow) -> row.Name) 
        |> generateBatchFile batchConfig 