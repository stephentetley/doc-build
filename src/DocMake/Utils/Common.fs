namespace DocMake.Utils

open System.IO
open System.Text

module Common = 

    let doubleQuote (s:string) : string = "\"" + s + "\""

    let safeName (input:string) : string = 
        let bads = ['\\'; '/'; ':']
        List.fold (fun s c -> s.Replace(c,'_')) input bads

    let zeroPad (width:int) (value:int) = 
        let ss = value.ToString ()
        let diff = width - ss.Length
        String.replicate diff "0" + ss

    let maybeCreateDirectory (dirpath:string) : unit = 
        if not <| Directory.Exists(dirpath)
        then 
            ignore <| Directory.CreateDirectory(dirpath)
        else 
            ()