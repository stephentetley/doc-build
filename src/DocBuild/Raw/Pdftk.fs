﻿// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Raw.Pdftk


// TODO - can use pdftk to get number of pages:
// > pdftk mydoc.pdf dump_data 
// Look for NumberOfPages in the output


[<AutoOpen>]
module Pdftk = 

    open DocBuild.Base.Common
    open DocBuild.Base.Shell

    

    let runPdftk (options:ProcessOptions) (command:string) : BuildResult<string> = 
        executeProcess options command