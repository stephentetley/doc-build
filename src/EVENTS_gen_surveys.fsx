#I @"..\packages\ExcelProvider.0.8.2\lib"
#r "ExcelProvider.dll"
open FSharp.ExcelProvider


#I @"..\packages\Newtonsoft.Json.10.0.3\lib\net45"
#r "Newtonsoft.Json"
open Newtonsoft.Json

// Need FAKE for @"DocMake\Base\Common.fs"
#I @"..\packages\FAKE.5.0.0-beta005\tools"
#r @"..\packages\FAKE.5.0.0-beta005\tools\FakeLib.dll"


#load @"DocMake\Base\Common.fs"
#load @"DocMake\Base\Json.fs"
#load @"DocMake\Base\GENHelper.fs"
open DocMake.Base.Common
open DocMake.Base.Json
open DocMake.Base.GENHelper


type SiteTable = 
    ExcelFile< @"G:\work\Projects\events2\EDM2 Site-List.xlsx",
               SheetName = "SITE_LIST",
               ForceString = false >

type SiteRow = SiteTable.Row

let siteTableDict : GetRowsDict<SiteTable, SiteRow> = 
    { GetRows     = fun imports -> imports.Data 
      NotNullProc = fun row -> match row.GetValue(0) with | null -> false | _ -> true }


let filterByBatch (batchName:string) (source:SiteRow list) : SiteRow list = 
    let matchBatch (row:SiteRow) : bool = 
        match row.``Survey Batch`` with
        | null -> false
        | ans -> ans = batchName
    List.filter  matchBatch source

let getSiteRows (batchName:string) : SiteRow list = 
    excelTableGetRows siteTableDict (new SiteTable()) |> filterByBatch batchName


let batchConfig : BatchFileConfig = 
    { PathToFake = @"D:\coding\fsharp\DocMake\packages\FAKE.5.0.0-beta005\tools\FAKE.exe"
      PathToScript = @"D:\coding\fsharp\DocMake\src\EVENTS_Survey_Sheets.fsx"
      OutputBatchFile = @"G:\work\Projects\events2\gen-surveys-risks\fake-make.bat"
      BuildTarget = "Final" }

let makeDict1 (row:SiteRow) : FindReplaceDict = 
    Map.ofList 
        <|  [ "#SITENAME",          row.Name
            ; "#SAINUMBER" ,        row.``SAI Number``
            ; "#SITEADDRESS",       row.``Site Address``
            ; "#OPERSTATUS",        row.``Operational Status``
            ; "#SITEGRIDREF",       row.``Site Grid Ref``
            ; "#ASSETTYPE",         row.Type
            ; "#OPERNAME",          row.``Operational Responsibility``
            ; "#OUTFALLGRIDREF",    row.``Outfall Grid Ref (from IW sheet)``
            ; "#RECWATERWOURSE",    row.``Receiving Watercourse``
            ]



let findReplaceConfig:FindsReplacesConfig<SiteRow> = 
    let makeFileName (row:SiteRow) = sprintf "%s_findreplace.json" (safeName row.Name)

    { DictionaryBuilder = makeDict1
      GetFileName = makeFileName
      OutputJsonFolder = @"G:\work\Projects\events2\gen-surveys-risks\__Json" }





// Generating all takes too long just generate a batch.
let main (batchName:string) : unit = 
    let siteList = getSiteRows batchName
    printfn "%i sites for output..." siteList.Length
    // Batch file
    siteList
        |> List.map (fun (row:SiteRow) -> row.Name) 
        |> generateBatchFile batchConfig 

    // Json
    siteList 
        |> Seq.iter (generateFindsReplacesJson findReplaceConfig)
