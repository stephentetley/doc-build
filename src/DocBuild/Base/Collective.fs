// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Base.Collective

[<AutoOpen>]
module Collective = 

    open DocBuild.Base.Document

    type Collective = 
        val private DocCollection : Document list

        new () = { DocCollection = [] }

        new (docs:Document list) = 
            { DocCollection = docs }

        new (paths:FilePath list) = 
            { DocCollection = 
                paths |> List.map (fun s -> Document(s)) }

        member x.Documents 
            with get () : Document list = x.DocCollection
            


