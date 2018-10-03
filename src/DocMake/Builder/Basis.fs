// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

module DocMake.Builder.Basis

open System.IO
open System.Threading



open DocMake.Base.Common
open DocMake.Base.FakeLike
open DocMake.Builder.BuildMonad
open DocMake.Builder.Document



let assertFile(fileName:string) : BuildMonad<'res,string> =  
    if File.Exists(fileName) then 
        breturn(fileName)
    else 
        throwError <| sprintf "assertFile failed: '%s'" fileName

// assertExtension ?

let getDocument (fileName:string) : BuildMonad<'res,Document<'a>> =   
    if File.Exists(fileName) then 
        breturn (makeDocument fileName)
    else 
        throwError <| sprintf "getDocument failed: '%s'" fileName


let copyToWorkingDirectory (fileName:string) : BuildMonad<'res,Document<'a>> = 
    if File.Exists(fileName) then 
        buildMonad { 
            let name1 = System.IO.FileInfo(fileName).Name
            let! cwd = asksEnv (fun e -> e.WorkingDirectory)
            let dest = System.IO.Path.Combine(cwd,name1)
            do System.IO.File.Copy(fileName,dest)
            return (makeDocument dest)
        }
    else 
        throwError <| sprintf "copyToWorkingDirectory failed: '%s'" fileName





let renameDocument (src:Document<'a>) (dest:string) : BuildMonad<'res,Document<'a>> = 
    match src.GetPath with
    | None -> breturn src
    | Some srcPath -> 
        executeIO <| fun () -> 
            let pathTo = System.IO.FileInfo(srcPath).DirectoryName
            let outPath = System.IO.Path.Combine(pathTo,dest)
            if System.IO.File.Exists(outPath) then System.IO.File.Delete(outPath)
            System.IO.File.Move(srcPath,outPath)
            makeDocument outPath

let renameTo (dest:string) (src:Document<'a>) : BuildMonad<'res,Document<'a>> = 
    renameDocument src dest


let askWorkingDirectory () : BuildMonad<'res,string> = 
    asksEnv (fun e -> e.WorkingDirectory)

let deleteWorkingDirectory () : BuildMonad<'res,unit> = 
    buildMonad { 
        let! cwd = askWorkingDirectory ()
        do printfn "Deleting: %s" cwd
        do! executeIO <| fun () -> deleteDirectory cwd
        do! executeIO <| fun () -> Thread.Sleep(360)
        }


let createWorkingDirectory () : BuildMonad<'res,unit> = 
    buildMonad { 
        let! cwd = asksEnv (fun e -> e.WorkingDirectory)
        do! executeIO (fun () -> maybeCreateDirectory cwd) 
    }

/// This should porbably be removed, it enables non-local writes...
let localWorkingDirectory (wd:string) (ma:BuildMonad<'res,'a>) : BuildMonad<'res,'a> = 
    localEnv (fun (e:Env) -> { e with WorkingDirectory = wd }) ma

let localSubDirectory (subdir:string) (ma:BuildMonad<'res,'a>) : BuildMonad<'res,'a> = 
    localEnv (fun (e:Env) -> 
                let cwd = System.IO.Path.Combine(e.WorkingDirectory, subdir)
                { e with WorkingDirectory = cwd }) 
            (createWorkingDirectory () >>. ma)


let shellRun (toolPath:string) (command:string)  (errMsg:string) : BuildMonad<'res, unit> = 
    printfn "----------\n%s %s\n----------\n" toolPath command
    try
        if 0 <> executeProcess toolPath command then
            throwError errMsg
        else            
            breturn ()
    with
    | ex -> 
        throwError (sprintf "shellRun: \n%s" ex.Message)


let makePdf (outputName:string) (proc:BuildMonad<'res, PdfDoc>) : BuildMonad<'res, PdfDoc> = 
    proc >>= renameTo outputName


let makeWordDoc (outputName:string) (proc:BuildMonad<'res, WordDoc>) :BuildMonad<'res, WordDoc> = 
    proc >>= renameTo outputName

