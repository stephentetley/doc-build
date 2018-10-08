// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

// This file contains functions and utilities to work with proprietry data 
// (usually spreadsheets).
// It is essentially "uninteresting".

module Proprietry

open FSharp.Data

open DocMake.Base.Common
open DocMake.Base.FakeLike
open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis




// *************************************
// Name helpers

let noParens (input:string) : string = 
    let parens = ['('; ')'; '['; ']'; '{'; '}']
    let ans1 = List.fold (fun (s:string) (c:char) -> s.Replace(c.ToString(), "")) input parens
    ans1.Trim() 

/// Does not replace parens, squares, braces...
let slashName (siteName:string) : string =
    siteName.Replace("_","/")


/// Does not replace parens, squares, braces...
let underscoreName (siteName:string) : string = 
    siteName.Replace("/","_")



// *************************************
// Upload spreadsheet



type SiteName = string
type SAINumber = string


// [<Literal>]
// let uploadSchema = @"Title(string), Sheet/Volume(string), Revision(string), Reference(string), Category(string), Project Name(string), File Format(string), File Date(string), Contractor(string)"
// let uploadHeaders = @"Title,Sheet/Volume,Revision,Reference,Category,Project Name,File Format,File Date,Contractor"

[<Literal>]
let UploadSchema = @"Asset Name(string), Asset Reference(string), Project Name(string), Project Code(string), Title(string), Category(string), Reference Number(string), Revision(string), Document Name(string), Document Date(string), Sheet/Volume(string)"
let uploadHeaders = @"Asset Name,Asset Reference,Project Name,Project Code,Title,Category,Reference Number,Revision,Document Name,Document Date,Sheet/Volume"



type UploadTable = 
    CsvProvider< Schema = UploadSchema,
                 HasHeaders = false >

type UploadRow = UploadTable.Row

let private outputUploadsCsv (source:seq<UploadRow>) (outputPath:string) : unit = 
    let table = new UploadTable(source)
    use sw = new System.IO.StreamWriter(outputPath)
    sw.WriteLine uploadHeaders
    table.Save(writer = sw, separator = ',', quote = '"' )



 
let standardDocumentDate () : string  = System.DateTime.Now.ToString("dd/MM/yyyy") 


type SiteRecord = 
    { SiteName: string 
      Uid: string }

/// F# design guidelines say favour object-interfaces rather than records of functions...
type IUploadHelper<'a> = 
    abstract member MakeUploadRow : SiteName -> SAINumber -> UploadRow
    abstract member ToSiteRecord : 'a -> SiteRecord


let makeUploadForm (helper:IUploadHelper<'site>) 
                    (siteRecords:'site list) : BuildMonad<'res,unit> = 
    buildMonad { 
        let saiDict = getSaiLookups ()
        let makeRow1 (osite:'site) : UploadRow = 
            let siteRec = helper.ToSiteRecord osite
            helper.MakeUploadRow siteRec.SiteName siteRec.Uid
        let rows = List.map makeRow1 siteRecords
        let! cwd = askWorkingDirectory () 
        let outPath = cwd </> "__EDMS_Upload.csv"
        do! executeIO (fun () -> outputUploadsCsv rows outPath)
    }       

