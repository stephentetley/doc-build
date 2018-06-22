// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


#load @"DocMake\Base\Common.fs"
#load @"DocMake\Base\FakeLike.fs"
#load @"DocMake\Tasks\PdfRotate.fs"
open DocMake.Base.Common
open DocMake.Tasks.PdfRotate


let test01 () = 
    let optsF = 
        fun (x:PdfRotateParams) ->
            { x with InputFile = @"G:\work\working\In1.pdf"
                     OutputFile = @"G:\work\working\Out1.pdf"
                     Rotations = [ (1,PoEast) ]  }
    PdfRotate optsF

