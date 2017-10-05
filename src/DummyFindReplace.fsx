
// Find Replace

#I @"C:\WINDOWS\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c"
#r "Microsoft.Office.Interop.Word"

#load @"DocMake\Utils\Common.fs"
#load @"DocMake\Utils\Office.fs"
#load @"DocMake\Tasks\DocFindReplace.fs"

open DocMake.Utils.Common
open DocMake.Tasks.DocFindReplace


let test01 () = printfn "%s" DocFindReplace_Teststring

let test02 () = safeName @"HAROLD/24:s01\z"
