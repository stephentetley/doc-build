// Office deps
#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.Word"
#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Excel\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.Excel"
#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.PowerPoint\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.PowerPoint"
#I @"C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c"
#r "office"

#I @"..\packages\Newtonsoft.Json.10.0.3\lib\net45"
#r "Newtonsoft.Json"
open Newtonsoft.Json

open System.IO

// FAKE reference path is local to the project file
#I @"..\packages\FAKE.5.0.0-beta005\tools"
#r @"..\packages\FAKE.5.0.0-beta005\tools\FakeLib.dll"
open Fake
open Fake.Core
open Fake.Core.Environment
open Fake.Core.Globbing.Operators
open Fake.Core.TargetOperators


#load @"DocMake\Base\Common.fs"
#load @"DocMake\Base\OfficeUtils.fs"
#load @"DocMake\Base\JsonUtils.fs"
#load @"DocMake\Tasks\DocFindReplace.fs"

open DocMake.Base.Common
open DocMake.Tasks.DocFindReplace



let _templateRoot   = @"G:\work\Projects\events2\gen-surveys-risks\__Templates"
let _jsonRoot       = @"G:\work\Projects\events2\gen-surveys-risks\__Json"
let _outputRoot     = @"G:\work\Projects\events2\gen-surveys-risks\output"


// siteName is an envVar so we can use this build script to build many 
// sites (they all follow the same directory/file structure).
let siteName = environVarOrDefault "sitename" @"MISSING"


let cleanName           = safeName siteName
let siteOutputDir       = _outputRoot @@ cleanName


let makeOutputName (fmt:Printf.StringFormat<string->string>) : string = 
    siteOutputDir @@ sprintf fmt cleanName

Target.Create "Clean" (fun _ -> 
    if Directory.Exists siteOutputDir then 
        Trace.tracefn " --- Clean folder: '%s' ---" siteOutputDir
        Fake.IO.Directory.delete siteOutputDir
    else 
        Trace.tracefn " --- Clean --- : folder does not exist '%s' ---" siteOutputDir
)

Target.Create "OutputDirectory" (fun _ -> 
    Trace.tracefn " --- Output folder: '%s' ---" siteOutputDir
    maybeCreateDirectory siteOutputDir
)

Target.Create "SurveySheet" (fun _ ->
    let template = _templateRoot @@ "TEMPLATE EDM2 Survey.docx"
    let jsonSource = _jsonRoot @@ (sprintf "%s_findreplace.json" cleanName)
    let docName = makeOutputName "%s EDM2 Survey.docx"
    
    DocFindReplace (fun p -> 
        { p with 
            TemplateFile = template
            OutputFile = docName
            JsonMatchesFile  = jsonSource 
        }) 
)

Target.Create "HazardSheet" (fun _ ->
    let template = _templateRoot @@ "TEMPLATE Hazard Identification Check List.docx"
    let jsonSource = _jsonRoot @@ (sprintf "%s_findreplace.json" cleanName)
    let docName = makeOutputName "%s Hazard Identification Check List.docx"
    
    DocFindReplace (fun p -> 
        { p with 
            TemplateFile = template
            OutputFile = docName
            JsonMatchesFile  = jsonSource 
        }) 
 
)

Target.Create "Final" (fun _ ->
    Trace.log "Done"
)

"Clean"
    ==> "OutputDirectory"

"OutputDirectory"
    ==> "SurveySheet"
    ==> "HazardSheet"
    ==> "Final"


// Note seemingly Fake files must end with this...
Target.RunOrDefault "None"
