
open System.IO
open System.Text.RegularExpressions

#load "ConcatPdfs.fs"
open ConcatPdfs


let InputPath = @"G:\work\usar\NWSC-Batch1\OUTPUT\BEN RHYDDING_STW"
let OutputName = @"Example1 Test.pdf"
let GSPath = @"C:\programs\gs\gs9.15\bin\gswin64c.exe"


let rec allFiles dir = 
    seq { for file in Directory.GetFiles dir do
            yield file
          for subdir in Directory.GetDirectories dir do
            yield! allFiles subdir }

let filterPDFs (ss:string seq) = 
    let re = new Regex("\.pdf$")
    Seq.filter (fun s -> re.Match(s).Success) ss




let pdfList () = allFiles InputPath |> filterPDFs |> Seq.iter (printfn "%s")

//let makeCmd () = 
//    let line1 = String.concat " " [ doubleQuote GSPath; GsOptions; outputOption OutputName ]
//    let linesK = allFiles InputPath |> filterPDFs |> Seq.map doubleQuote
//    String.concat "^\n" (line1 :: Seq.toList linesK)

let searches = [ "MM3x.*\.pdf$"
               ; "Site Works.*\.pdf$" 
               ; "Install Photos.*\.pdf$" ]

let test01 () = 
    let ins = makeInputList InputPath searches
    ins |> List.iter (fun s -> printfn "%s" s)


let test02 () = 
    let ins = makeInputList InputPath searches
    let bat = makeCmd GSPath ins "TEMP1.pdf"
    File.WriteAllText(@"G:\work\usar\NWSC-Batch1\OUTPUT\BEN RHYDDING_STW\mk.bat", bat)

let makePdfName (infolder:DirectoryInfo) : string = 
    sprintf "%s - AMP6 Ultrasonic Asset Replacement Manual.pdf" infolder.Name

let RunMain () = 
    processMany GSPath @"G:\work\rtu\mk3-mk4-upgrades" searches