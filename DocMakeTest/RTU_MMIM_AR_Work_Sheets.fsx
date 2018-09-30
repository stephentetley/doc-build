// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


// Office deps
#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.Word"
#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Excel\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.Excel"
#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.PowerPoint\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.PowerPoint"
#I @"C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c"
#r "office"

// Use FSharp.Data for CSV output
#I @"..\packages\FSharp.Data.3.0.0-beta3\lib\net45"
#r @"FSharp.Data.dll"
open FSharp.Data


#I @"..\packages\ExcelProvider.1.0.1\lib\net45"
#r "ExcelProvider.Runtime.dll"

#I @"..\packages\ExcelProvider.1.0.1\typeproviders\fsharp41\net45"
#r "ExcelDataReader.DataSet.dll"
#r "ExcelDataReader.dll"
#r "ExcelProvider.DesignTime.dll"
open FSharp.Interop.Excel


#I @"..\packages\Magick.NET-Q8-AnyCPU.7.4.6\lib\net40"
#r @"Magick.NET-Q8-AnyCPU.dll"
open ImageMagick




#load "..\DocMake\DocMake\Base\Common.fs"
#load "..\DocMake\DocMake\Base\FakeLike.fs"
#load "..\DocMake\DocMake\Base\ExcelProviderHelper.fs"
#load "..\DocMake\DocMake\Base\OfficeUtils.fs"
#load "..\DocMake\DocMake\Base\ImageMagickUtils.fs"
#load "..\DocMake\DocMake\Base\SimpleDocOutput.fs"
open DocMake.Base.Common
open DocMake.Base.FakeLike
open DocMake.Base.ExcelProviderHelper

#load "..\DocMake\DocMake\Builder\BuildMonad.fs"
#load "..\DocMake\DocMake\Builder\Document.fs"
#load "..\DocMake\DocMake\Builder\Basis.fs"
#load "..\DocMake\DocMake\Builder\ShellHooks.fs"
open DocMake.Base.Common
open DocMake.Base.FakeLike
open DocMake.Builder.BuildMonad
open DocMake.Builder.Basis



#load "..\DocMake\DocMake\Base\Common.fs"
#load "..\DocMake\DocMake\Base\FakeLike.fs"
#load "..\DocMake\DocMake\Base\OfficeUtils.fs"
#load "..\DocMake\DocMake\Base\SimpleDocOutput.fs"
#load "..\DocMake\DocMake\Builder\BuildMonad.fs"
#load "..\DocMake\DocMake\Builder\Document.fs"
#load "..\DocMake\DocMake\Builder\Basis.fs"
#load "..\DocMake\DocMake\Tasks\DocFindReplace.fs"
#load "..\DocMake\DocMake\WordBuilder.fs"
open DocMake.Base.Common
open DocMake.Base.FakeLike
open DocMake.Base.OfficeUtils
open DocMake.Builder.BuildMonad
open DocMake.Builder.Document
open DocMake.Builder.Basis
open DocMake.WordBuilder

#load @"Proprietry.fs"
open Proprietry


let _templateRoot   = @"G:\work\Projects\rtu\MK5 MMIM Replacement\forms\__Templates"
let _outputRoot     = @"G:\work\Projects\rtu\MK5 MMIM Replacement\forms\output"

// Note - suffixing the sheet name with "!" seems to be needed
// when the required sheet is not the first sheet.
type SiteTable = 
    ExcelFile< @"G:\work\Projects\rtu\MK5 MMIM Replacement\SiteList-2010-2011-2012.xlsx",
               SheetName = "Sites_2010!",
               ForceString = true >

type SiteRow = SiteTable.Row


let readSiteRows () : SiteRow list = 
    let helper = 
        { new IExcelProviderHelper<SiteTable,SiteRow>
          with member this.ReadTableRows table = table.Data 
               member this.IsBlankRow row = match row.GetValue(0) with null -> true | _ -> false }
         
    excelReadRowsAsList helper (new SiteTable())


let makeSurveyMatches (row:SiteRow) : SearchList = 
    let getString (str:string) : string = 
        match str with
        | null -> ""
        | _ -> str

    [ "#SITENAME",          getString <| row.``Site Name``
    ; "#SAINUMBER" ,        getString <| row.``Sai Number``
    ; "#SITE_ADDRESS",      getString <| row.``Site Address``
    ; "#GRIDREF",           getString <| row.``Grid Ref``
    ; "#WORK_CENTRE",       getString <| row.``Work Centre``
    ; "#OPCONTACT",         getString <| row.``Operational Contact``
    ; "#RTS_OUTSTATION",    getString <| row.``Outstation Name``
    ; "#OUTSTATION_ADDR",   getString <| row.``RTU Address``
    ]

let makeHazardsMatches (site:SiteRow) : SearchList = 
    [ "#SITENAME",          site.``Site Name``
    ; "#SAINUMBER" ,        site.``Sai Number``
    ]

let genHazardSheet (row:SiteRow) : WordBuild<WordDoc> =
    buildMonad { 
        let templatePath = _templateRoot </> "TEMPLATE Hazard Identification Check List.docx"
        let cleanName = underscoreName row.``Site Name``
        let outName = sprintf "%s Hazard Identification Check List.docx" cleanName   
        let matches = makeHazardsMatches row
        let! template = getTemplateDoc templatePath
        let! d1 = docFindReplace matches template >>= renameTo outName 
        return d1
    }

let genSurvey (row:SiteRow) : WordBuild<WordDoc> = 
    buildMonad { 
        let surveyTemplate = _templateRoot </> "Expired MMIM Form.docx"
        let cleanName = underscoreName row.``Site Name``
        let outName = sprintf "%s MMIM Upgrade Site Words.docx" cleanName   
        let matches = makeSurveyMatches row
        let! template = getTemplateDoc surveyTemplate
        let! d1 = docFindReplace matches template >>= renameTo outName 
        return d1
    }

let build1 (row:SiteRow) : WordBuild<unit> = 
    let dirName = underscoreName row.``Site Name``
    localSubDirectory dirName <| 
        (genHazardSheet row >>. genSurvey row >>. breturn ())
            
let buildScript () : WordBuild<unit> = 
    let siteRows = readSiteRows ()
    mapMz build1 siteRows

        

let main () : unit = 
    let env = 
        { WorkingDirectory = _outputRoot
          PrintQuality = PrintQuality.PqScreen
          PdfQuality = PdfPrintQuality.PdfScreen }

    runWordBuild env (buildScript ()) 

