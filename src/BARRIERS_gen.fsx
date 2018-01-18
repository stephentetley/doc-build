// Only Generate Batch file 
// Find/Replace not needed

#I @"..\packages\ExcelProvider.0.8.2\lib"
#r "ExcelProvider.dll"
open FSharp.ExcelProvider


#I @"..\packages\Newtonsoft.Json.10.0.3\lib\net45"
#r "Newtonsoft.Json"
open Newtonsoft.Json

#load @"DocMake\Base\Json.fs"
#load "GENHelper.fs"
open GENHelper


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
      OutputBatchFile = @"G:\work\Projects\barriers\fake-make.bat" }

let main () : unit = 
    getSitesRows () 
        |> List.map (fun (row:SitesRow) -> row.Name) 
        |> generateBatchFile batchConfig 