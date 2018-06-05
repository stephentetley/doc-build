// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

#I @"..\packages\ExcelProvider.0.8.2\lib"
#r "ExcelProvider.dll"
open FSharp.ExcelProvider

#I @"..\packages\Magick.NET-Q8-AnyCPU.7.4.6\lib\net40"
#r "Newtonsoft.Json"
open Newtonsoft.Json

// FAKE is local to the project file
#I @"..\packages\FAKE.5.0.0-rc016.225\tools"
#r @"FakeLib.dll"
#I @"..\packages\Fake.Core.Globbing.5.0.0-beta021\lib\net46"
#r @"Fake.Core.Globbing.dll"
#I @"..\packages\Fake.IO.FileSystem.5.0.0-rc017.237\lib\net46"
#r @"Fake.IO.FileSystem.dll"
#I @"..\packages\Fake.Core.Trace.5.0.0-rc017.237\lib\net46"
#r @"Fake.Core.Trace.dll"
#I @"..\packages\Fake.Core.Process.5.0.0-rc017.237\lib\net46"
#r @"Fake.Core.Process.dll"

open System

#load @"DocMake\Base\Common.fs"
#load @"DocMake\Base\JsonUtils.fs"
#load @"DocMake\Base\GENHelper.fs"
open DocMake.Base.Common
open DocMake.Base.JsonUtils
open DocMake.Base.GENHelper

type SiteTable = 
    ExcelFile< @"G:\work\Projects\flow2\final-docs\Input\Batch02\sitelist.xlsx",
               SheetName = "Sheet1",
               ForceString = false >

type SiteRow = SiteTable.Row

let siteTableDict : ExcelProviderHelperDict<SiteTable, SiteRow> = 
    { GetRows     = fun imports -> imports.Data 
      NotNullProc = fun row -> match row.GetValue(0) with | null -> false | _ -> true }

let getSitesRows () : SiteRow list = excelTableGetRows siteTableDict (new SiteTable())



let makeDict (row:SiteRow) : FindReplaceDict = 
    Map.ofList [ "#SITENAME", row.Site
               ; "#SAINUM" , row.Uid
               ]

let jsonConfig : FindsReplacesConfig<SiteRow> = 
    { DictionaryBuilder = makeDict
    ; GetFileName       = 
        fun (row:SiteRow) -> sprintf "%s_findreplace.json" (safeName row.Site)

    ; OutputJsonFolder = @"G:\work\Projects\flow2\final-docs\__Json" }

let batchConfig : BatchFileConfig = 
    { PathToFake = @"D:\coding\fsharp\DocMake\packages\FAKE.5.0.0-beta005\tools\FAKE.exe"
      PathToScript = @"D:\coding\fsharp\DocMake\src\FC2_Final_Build.fsx"
      BuildTarget = "Final"
      OutputBatchFile = @"G:\work\Projects\flow2\final-docs\fake-make.bat"
      VarName = "sitename" }


// A file is generated foreach row
let main () : unit = 
    let siteList = getSitesRows () 
    // Generate find-replace json...
    siteList |> List.iter (generateFindsReplacesJson jsonConfig)
    // Generate batch file...
    siteList 
        |> List.map (fun (row:SiteRow) -> row.Site) 
        |> generateBatchFile batchConfig 