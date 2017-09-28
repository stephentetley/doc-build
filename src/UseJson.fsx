
#I @"..\packages\FSharp.Data.2.4.0\lib\net40"
#r "FSharp.Data.dll"

#r "System.Globalization"
#r "System.IO"

open System.Globalization
open System.IO
open FSharp.Data
open FSharp.Data.JsonExtensions

let info = 
    JsonValue.Parse(""" 
        { "siteName": "AUBURN/WTW", "saiNumber": "SAI067" }
        """)


let name = info?siteName.AsString ()

let demo02 () =
    printfn "%s" (__SOURCE_DIRECTORY__ + @"\..\data\site1.json")
    use input = new StreamReader ( __SOURCE_DIRECTORY__ + @"\..\data\site1.json")
    let info2 = JsonValue.Parse(input.ReadToEnd() )
    printfn "%s:%s" (info2?siteName.AsString ()) (info2?saiNumber.AsString ())
