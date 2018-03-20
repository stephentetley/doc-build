module DocMake.Base.GENHelper

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
      BuildTarget : string
      OutputBatchFile : string }



let private doubleQuote (s:string) : string = sprintf "\"%s\"" s

type SiteName = string

let private genInvoke1 (sw:IO.StreamWriter) (config:BatchFileConfig) (count:int) (ix:int) (siteName:string) : unit = 
    fprintf sw "REM %s (%d of %d) ... \n"  siteName (ix+1) count
    fprintf sw "%s ^\n"  (doubleQuote config.PathToFake)
    fprintf sw "    %s ^\n"  (doubleQuote config.PathToScript)
    fprintf sw "    %s --envar sitename=%s\n\n" config.BuildTarget (doubleQuote siteName)

let generateBatchFile (config:BatchFileConfig) (siteNames:string list) : unit = 
    let count = List.length siteNames
    use sw = new IO.StreamWriter(config.OutputBatchFile)
    fprintf sw "@echo on\n\n"
    List.iteri (genInvoke1 sw config count) siteNames
    sw.Close ()


// Generate JSON files of name/values pairs for running Find/Replace 
// on Word docs.

type FindsReplacesConfig<'row> = 
    { DictionaryBuilder : 'row -> FindReplaceDict
      GetFileName : 'row -> string
      OutputJsonFolder : string }

let generateFindsReplacesJson (config:FindsReplacesConfig<'row>) (row:'row) : unit = 
    let file = config.GetFileName row
    let fullName = System.IO.Path.Combine(config.OutputJsonFolder, file)
    let dict:FindReplaceDict = config.DictionaryBuilder row
    writeJsonFindReplaceDict fullName dict


