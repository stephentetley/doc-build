[<AutoOpen>]
module Tasks.DocToPdf

open Microsoft.Office.Interop.Word

open System.IO
open System.Text.RegularExpressions

let refobj (x:'a) : ref<obj> = ref (x :> obj)


let outputFolderName (inpath:string) (outroot:string) : string =
    let leaf = if Directory.Exists(inpath) then DirectoryInfo(inpath) else DirectoryInfo(inpath).Parent
    System.IO.Path.Combine(outroot, leaf.Name)

let outputFileName (inpath:string) (outroot:string) : string =
    let leaf = DirectoryInfo(inpath)
    let leaf1 = DirectoryInfo(inpath).Parent.Name
    let leaf2 = System.IO.Path.ChangeExtension(leaf.Name, "pdf")
    System.IO.Path.Combine(outroot, leaf1, leaf2)

let matchingFiles (folder:string) (re:Regex) : string seq = 
    let kids = Directory.GetFiles(folder)
    kids |> Seq.filter (fun a -> re.IsMatch(a))

let process1 (app:Application) (inpath:string) (outroot:string) : unit = 
    let outfolder = outputFolderName inpath outroot
    let outfile = outputFileName inpath outroot
    printfn "%s ==> %s" inpath outfolder
    if Directory.Exists(outfolder) then () else (ignore <| Directory.CreateDirectory (outfolder))
    try 
        let doc = app.Documents.Open(FileName = refobj inpath)
        doc.ExportAsFixedFormat (OutputFileName = outfile, ExportFormat = WdExportFormat.wdExportFormatPDF)
        doc.Close (SaveChanges = refobj false)
    with
    | ex -> printfn "Some error occured - %s - %s" inpath ex.Message


let processFolder (app:Application) (infolder:DirectoryInfo) (outroot:string) : unit = 
    // Does not contain "~$"; ends with suffix ".docx"
    let re = new Regex ("\.docx$")
    let re2 = new Regex ("^((?!~\$).)*$")
    let docs = matchingFiles infolder.FullName re
    docs |> Seq.filter (fun a -> re2.IsMatch(a))
         |> Seq.iter (fun a -> process1 app a outroot)
    

let processMany (inroot:string) (outroot:string) : unit = 
    let app = new ApplicationClass (Visible = false)
    match Directory.Exists inroot with
    | false -> printfn "Missing dir: %s" inroot
    | true -> let app = new ApplicationClass (Visible = false)
              let indir = DirectoryInfo(inroot)
              indir.GetDirectories () |> Seq.iter (fun dir -> processFolder app dir outroot)
              app.Quit ()
              printfn "Outroot: %s" outroot

