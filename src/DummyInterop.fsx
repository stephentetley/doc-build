// Executing this in F# Interactive should give the full path of 'Microsoft.Office.Interop.Word.dll'
// (it is located in the GAC - Global Assembly Cache).
//
// We can then (temporarily) use the GAC path in build scripts though this is entirely non-portable, 
// and we really need a proper soultion.

// Note - "Microsoft Office **.* Object Library" is called "office" in the GAC

#r "Microsoft.Office.Interop.Word"
#r "Microsoft.Office.Interop.Excel"
#r "Microsoft.Office.Interop.PowerPoint"
#r "office"

let main () = printfn "hello"
