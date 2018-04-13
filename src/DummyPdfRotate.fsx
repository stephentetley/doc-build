// FAKE is local to the project file
#I @"..\packages\FAKE.5.0.0-beta005\tools"
#r @"..\packages\FAKE.5.0.0-beta005\tools\FakeLib.dll"

#load @"DocMake\Base\Common.fs"
#load @"DocMake\Base\FakeExtras.fs"
#load @"DocMake\Tasks\PdfRotate.fs"
open DocMake.Base.Common
open DocMake.Base.FakeExtras
open DocMake.Tasks.PdfRotate


let test01 () = 
    let optsF = 
        fun (x:PdfRotateParams) ->
            { x with InputFile = @"G:\work\working\In1.pdf"
                     OutputFile = @"G:\work\working\Out1.pdf"
                     Rotations = [ (1,PoEast) ]  }
    PdfRotate optsF

