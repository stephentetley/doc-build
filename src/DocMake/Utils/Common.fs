namespace DocMake.Utils

open System.Text

module Common = 

    let doubleQuote (s:string) : string = "\"" + s + "\""

    let safeName (input:string) : string = 
        let bads = ['\\'; '/'; ':']
        List.fold (fun s c -> s.Replace(c,'_')) input bads

