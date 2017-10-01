[<AutoOpen>]
module DocMake.Tasks.PdfConcat

// Concat PDFs

open System.IO
open System.Text.RegularExpressions
open DocMake.Utils.Common

[<CLIMutable>]
type PdfConcatParams = 
    { 
      Output : string
      ToolPath : string
      GsOptions : string
    }

// Note input files are supplied as arguments to the top level "Command".
// e.g CscHelper.fs
// Csc : (CscParams -> CscParams) * string list -> unit



// ArchiveHelper.fs includes output file name as a function argument
// DotCover.fs includes output file name in the params
// FscHelper includes output file name in the params

let PdfConcatDefaults = 
    { Output = "concat.pdf"
      ToolPath = "C:\programs\gs\gs9.15\bin\gswin64c.exe"
      GsOptions = @"-dBATCH -dNOPAUSE -q -sDEVICE=pdfwrite -dPDFSETTINGS=/prepress" }


let line1 (opts:PdfConcatParams) : string =
    sprintf "%s %s -sOutputFile=%s" (doubleQuote opts.ToolPath) opts.GsOptions (doubleQuote opts.Output)

let lineK (name:string) : string = sprintf " \"%s\"" name

let unlines (lines: string list) : string = String.concat "\n" lines
let unlinesC (lines: string list) : string = String.concat "^\n" lines

let makeCmd (setPdfConcatParams: PdfConcatParams -> PdfConcatParams) (inputFiles: string list) : string = 
    let parameters = PdfConcatDefaults |> setPdfConcatParams
    let first = line1 parameters
    let rest = List.map lineK inputFiles
    unlinesC <| first :: rest

// TODO run as a process...


// Below is old...


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
