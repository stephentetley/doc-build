// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Base.Collective

[<AutoOpen>]
module Collective = 

    open DocBuild.Base.Document

    type Collective = 
        val private Documents : Document list

        new () = { Documents = [] }

        new (docs:Document list) = 
            { Documents = docs }

        new (paths:FilePath list) = 
            { Documents = 
                paths |> List.map (fun s -> new Document(filePath = s)) }


