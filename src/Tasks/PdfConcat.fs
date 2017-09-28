[<AutoOpen>]
module Tasks.PdfConcat

// Concat PDFs

open System.IO
open System.Text.RegularExpressions

let GsOptions = @"-dBATCH -dNOPAUSE -q -sDEVICE=pdfwrite -dPDFSETTINGS=/prepress"


let doubleQuote s = "\"" + s + "\""

let outputFileName (inpath:string) : string =
    let pathto = if Directory.Exists(inpath) then DirectoryInfo(inpath) else DirectoryInfo(inpath).Parent
    System.IO.Path.Combine(pathto.FullName, "make-final.bat")


let matchingFiles (folder:string) (re:Regex) : string list = 
    let kids = Directory.GetFiles(folder)
    kids |> Seq.filter (fun a -> re.IsMatch(a)) |> Seq.toList

// Need to make an ordered list of matching files - find with a Regex...
let makeInputList (folder:string) (regexs : string list) : string list =
    let mkRegex = fun ss -> new Regex(ss)
    List.concat <| List.map (fun re -> matchingFiles folder (mkRegex re)) regexs

let outputOption (name:string) = "-sOutputFile=" + doubleQuote name

// TODO - this should be lifted out to user code
let makePdfName (infolder:DirectoryInfo) : string = 
    sprintf "%s AMP6 RTU MK3_MK4 Upgrade Manual.pdf" infolder.Name

let makeCmd (gspath:string) (infiles:string list) (outfile:string) :string  = 
    let line1 = String.concat " " [ doubleQuote gspath; GsOptions; outputOption outfile ]
    let linesK = List.map (fun s -> " " + doubleQuote s) infiles
    String.concat "^\n" (line1 :: Seq.toList linesK)

let processFolder (gspath:string) (dir:DirectoryInfo) (regexs : string list): unit = 
    let pdfname = makePdfName dir
    let ins = makeInputList dir.FullName regexs
    let cmd = makeCmd gspath ins pdfname
    let outpath = outputFileName dir.FullName
    File.WriteAllText(outpath, cmd)
    ()

let processMany (gspath:string) (inroot:string) (regexs : string list): unit = 
    match Directory.Exists inroot with
    | false -> printfn "Missing dir: %s" inroot
    | true -> let indir = DirectoryInfo(inroot)
              indir.GetDirectories () |> Seq.iter (fun dir -> processFolder gspath dir regexs)
