// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Base

[<AutoOpen>]
module BaseDefinitions = 


    type PageOrientation = 
        | Portrait 
        | Landscape



    type PrintQuality = 
        | Screen
        | Print


    // ************************************************************************
    // Find and replace


    type SearchList = (string * string) list