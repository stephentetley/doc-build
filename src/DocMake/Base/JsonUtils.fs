module DocMake.Base.JsonUtils

open System

open Newtonsoft.Json

open DocMake.Base.Common


let readJsonFindReplaceDict (jsonFile:string) : FindReplaceDict = 
    let readProc (json:string) : FindReplaceDict =  JsonConvert.DeserializeObject<FindReplaceDict>(json) 
    use sr = new IO.StreamReader(jsonFile)
    readProc <| sr.ReadToEnd()

let readJsonStringPairs (jsonFile:string) : (string * string) list = 
    readJsonFindReplaceDict jsonFile |> Map.toList

let writeJsonFindReplaceDict (jsonFile:string) (dict:FindReplaceDict) : unit = 
    let write1 (name:string) (value:string) (w:JsonTextWriter) : unit = 
        w.WritePropertyName name
        w.WriteValue value
    use sw : System.IO.StreamWriter = new IO.StreamWriter(path=jsonFile,append=false)
    use handle : JsonTextWriter = new JsonTextWriter(sw)
    handle.Formatting <- Formatting.Indented
    handle.Indentation <- 2
    handle.WriteStartObject ()
    Map.iter (fun k v  -> write1 k v handle) dict    
    handle.WriteEndObject ()

let writeJsonStringPairs (jsonFile:string) (pairs: (string * string) list) : unit = 
    writeJsonFindReplaceDict jsonFile <| Map.ofList pairs

