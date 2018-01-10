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


type SitesTable = 
    ExcelFile< @"G:\work\Projects\samps\sitelist-for-gen-jan2018.xlsx",
               SheetName = "Sheet1",
               ForceString = false >

type SitesRow = SitesTable.Row

let jsonFolder = @"G:\work\Projects\samps\Final_Docs\__Json"


let makeDict (row:SitesRow) : Dict = 
    Map.ofList [ "#SITENAME", row.Site
               ; "#SAINUM" , row.Uid
               ]

let processRow (row:SitesRow) : unit = 
    let name1 = sprintf "%s_findreplace.json" (safeName row.Site)
    let fileName = System.IO.Path.Combine(jsonFolder, name1)
    let dict = makeDict row
    writeJsonDict fileName dict

let main () : unit = 
    let masterData = new SitesTable()
    let nullPred (row:SitesRow) = match row.Site with null -> false | _ -> true

    masterData.Data 
        |> Seq.filter nullPred
        |> Seq.iter processRow