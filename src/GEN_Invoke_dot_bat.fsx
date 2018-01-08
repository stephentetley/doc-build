#I @"..\packages\ExcelProvider.0.8.2\lib"
#r "ExcelProvider.dll"
open FSharp.ExcelProvider

open System


type SitesTable = 
    ExcelFile< @"G:\work\Projects\samps\sitelist-for-gen-jan2018.xlsx",
               SheetName = "Sheet1",
               ForceString = false >

type SitesRow = SitesTable.Row

let pathToFake = @"D:\coding\fsharp\DocMake\packages\FAKE.5.0.0-beta005\tools\FAKE.exe"
let pathToScript = @"D:\coding\fsharp\DocMake\src\SAMPS_Final_Build.fsx"
let outputBat = @"G:\work\Projects\samps\fake-make.bat"

let doubleQuote (s:string) : string = sprintf "\"%s\"" s

let genInvoke1 (sw:IO.StreamWriter) (row:SitesRow) : unit = 
    fprintf sw "REM %s ...\n"  row.Site
    fprintf sw "%s ^\n"  (doubleQuote pathToFake)
    fprintf sw "    %s ^\n"  (doubleQuote pathToScript)
    fprintf sw "    Final --envar sitename=%s --envar uid=%s\n\n"
               (doubleQuote row.Site)
               (doubleQuote row.Uid)

let main () : unit = 
    let masterData = new SitesTable()
    let nullPred (row:SitesRow) = match row.Site with null -> false | _ -> true
    use sw = new IO.StreamWriter(outputBat)
    fprintf sw "@echo off\n\n"
    masterData.Data 
        |> Seq.filter nullPred
        |> Seq.iter (genInvoke1 sw)
    sw.Close ()