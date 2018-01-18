module GENHelper

open System

open DocMake.Base.Json

// The Excel Type Provider seems to read a trailing null row.
// This dictionary and procedure provide a skeleton to get round this.

type GetRowsDict<'table, 'row> = 
    { GetRows : 'table -> seq<'row>
      NotNullProc : 'row -> bool }

let excelTableGetRows (dict:GetRowsDict<'table,'row>) (table:'table) : 'row list = 
    let allrows = dict.GetRows table
    allrows |> Seq.filter dict.NotNullProc |> Seq.toList


// Generate Batch file

type BatchFileConfig = 
    { PathToFake : string
      PathToScript : string
      OutputBatchFile : string }



let private doubleQuote (s:string) : string = sprintf "\"%s\"" s

type SiteName = string

let private genInvoke1 (sw:IO.StreamWriter) (config:BatchFileConfig) (siteName:string) : unit = 
    fprintf sw "REM %s ...\n"  siteName
    fprintf sw "%s ^\n"  (doubleQuote config.PathToFake)
    fprintf sw "    %s ^\n"  (doubleQuote config.PathToScript)
    fprintf sw "    Final --envar sitename=%s\n\n"  (doubleQuote siteName)

let generateBatchFile (config:BatchFileConfig) (siteNames:string list) : unit = 
    use sw = new IO.StreamWriter(config.OutputBatchFile)
    fprintf sw "@echo off\n\n"
    List.iter (genInvoke1 sw config) siteNames
    sw.Close ()


// Generate JSON files of name/values pairs for string matching



type FindsReplacesConfig<'row> = 
    { DictionaryBuilder : 'row -> Dict
      GetFileName : 'row -> string
      OutputJsonFolder : string }

let generateFindsReplacesJson (config:FindsReplacesConfig<'row>) (row:'row) : unit = 
    let file = config.GetFileName row
    let fullName = System.IO.Path.Combine(config.OutputJsonFolder, file)
    let dict:Dict = config.DictionaryBuilder row
    writeJsonDict fullName dict


