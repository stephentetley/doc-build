namespace DocMake.Base

open System

open Newtonsoft.Json


module Json = 

    type Dict = Map<string,string>

    let readJsonDict (jsonFile:string) : Dict = 
        let readProc (json:string) : Dict =  JsonConvert.DeserializeObject<Dict>(json) 
        use sr = new IO.StreamReader(jsonFile)
        readProc <| sr.ReadToEnd()

    let readJsonStringPairs (jsonFile:string) : (string * string) list = 
        readJsonDict jsonFile |> Map.toList

    let writeJsonDict (jsonFile:string) (dict:Dict) : unit = 
        let write1 (name:string) (value:string) (w:JsonTextWriter) : unit = 
            w.WritePropertyName name
            w.WriteValue value
        use sw : System.IO.StreamWriter = new IO.StreamWriter(jsonFile)
        use handle : JsonTextWriter = new JsonTextWriter(sw)
        handle.WriteStartObject ()
        Map.iter (fun k v  -> write1 k v handle) dict    
        handle.WriteEndObject ()

    let writeJsonStringPairs (jsonFile:string) (pairs: (string * string) list) : unit = 
        writeJsonDict jsonFile <| Map.ofList pairs

