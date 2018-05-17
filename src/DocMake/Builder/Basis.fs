module DocMake.Builder.Basis

open System.IO

open DocMake.Builder.BuildMonad



let assertFile(fileName:string) : BuildMonad<'res,string> =  
    if File.Exists(fileName) then 
        buildMonad.Return(fileName)
    else 
        throwError <| sprintf "assertFile failed: '%s'" fileName

// assertExtension ?