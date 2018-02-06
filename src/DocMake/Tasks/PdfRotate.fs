module DocMake.Tasks.PdfRotate

open System.IO

open Fake.Core
open Fake.Core.Process

open DocMake.Base.Common

type PageRotation = int * DocMakePageOrientation

[<CLIMutable>]
type PdfRotateParams = 
    { InputFile: string
      OutputFile: string
      PdftkExePath: string
      Rotations: PageRotation list }


let PdfRotateDefaults = 
    { InputFile = @""
      OutputFile = "rotate_out.pdf"
      PdftkExePath = @"C:\programs\PDFtk Server\bin\pdftk.exe" 
      Rotations = [] }


let private makeRotateSpec (rotations: PageRotation list) : string = 
    let rec work inlist start ac = 
        match inlist with
        | [] -> 
            let final = sprintf "A%i-end" start
            List.rev (final::ac)
        | (pageNum,po) :: rest -> 
            if pageNum = start then 
                let thisRotation = sprintf "A%i%s" pageNum (pdftkPageOrientation po)
                work rest (pageNum+1) (thisRotation :: ac)
            else 
                let thisRotation = sprintf "A%i%s" pageNum (pdftkPageOrientation po)
                let rangeAsIs = sprintf "A%i-%i" start (pageNum-1)
                work rest (pageNum+1) (thisRotation :: rangeAsIs :: ac)
    String.concat " " <| work rotations 1 []

let private makeCmd (parameters: PdfRotateParams)  : string = 
    let rotateSpec = "cat A1east A2-end" // temp
    sprintf "A=\"%s\" %s output \"%s\"" parameters.InputFile rotateSpec parameters.OutputFile 


// Run as a process...
let private shellRun toolPath command = 
    printfn "%s %s" toolPath command
    if 0 <> ExecProcess (fun info -> 
                info.FileName <- toolPath
                info.Arguments <- command) System.TimeSpan.MaxValue
    then failwithf "PdfRotate %s failed." command


let PdfRotate (setPdfRotateParams: PdfRotateParams -> PdfRotateParams)  : unit =
    let parameters = PdfRotateDefaults |> setPdfRotateParams
    let command = makeCmd parameters
    shellRun parameters.PdftkExePath command
