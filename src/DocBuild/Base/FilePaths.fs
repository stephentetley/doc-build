// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause

namespace DocBuild.Base

[<AutoOpen>]
module FilePaths = 

    open System.IO

    type PathSegments = string list


    let private directoryStep (directory:DirectoryInfo) (initialAcc:string list) : string list =
        let rootName = directory.Root.Name
        let rec work (currentDir:DirectoryInfo) (acc:string list) = 
            let folderName = currentDir.Name
            if folderName = rootName then
                rootName :: acc
            else
                work currentDir.Parent (folderName :: acc)
        work directory initialAcc
            

    let private filePathSegments (path:string) = 
        let fileInfo = new FileInfo(path)
        directoryStep fileInfo.Directory [fileInfo.Name]

    let private directoryPathSegments (path:string) = 
        directoryStep (new DirectoryInfo(path)) []

    let private localPath (segments:string list) = 
        let sep = Path.DirectorySeparatorChar.ToString()
        match segments with
        | root :: rest -> 
            if root.EndsWith(sep) then 
                root + String.concat sep rest
            else
                String.concat sep segments
        | [] -> ""

    
    // ************************************************************************
    // Abstract over FilePaths and DirectoryPaths

    type HasPathSegments =
        abstract GetPathSegments : string list
        

    [<Struct>]
    type FilePath = 
        | FilePath of string

        member private x.Body 
            with get () : string = match x with | FilePath(s) -> s

        member x.Segments
            with get () : string list = filePathSegments x.Body

        member x.LocalPath 
            with get () : string = localPath x.Segments

        interface HasPathSegments with
            member x.GetPathSegments = x.Segments

    [<Struct>]
    type DirectoryPath = 
        | DirectoryPath of string

        member private x.Body 
            with get () : string = match x with | DirectoryPath(s) -> s

        member x.Segments
            with get () : string list = directoryPathSegments x.Body

        member x.LocalPath 
            with get () : string = localPath x.Segments

        interface HasPathSegments with
            member x.GetPathSegments = x.Segments

    let commonPathPrefix (x:#HasPathSegments) (y:#HasPathSegments) : string = 
        let rec work xs ys (acc:string list) = 
            match xs, ys with
            | (s :: ss), (t :: ts) -> 
                if s = t then 
                    work ss ts (s :: acc)
                else
                    List.rev acc
            | _, _ -> List.rev acc
        work x.GetPathSegments y.GetPathSegments [] |> localPath



    let rightPathComplement (root:DirectoryPath) (y:#HasPathSegments) : string = 
        let rec work xs ys  = 
            match xs, ys with
            | (s :: ss), (t :: ts) -> 
                if s = t then 
                    work ss ts
                else ys
            | _, _ -> ys
        work root.Segments y.GetPathSegments  |> localPath


    let rootIsPrefix (root:DirectoryPath) (y:#HasPathSegments) : bool = 
        let rec work xs ys  = 
            match xs, ys with
            | (s :: ss), (t :: ts) -> 
                if s = t then 
                    work ss ts
                else false
            | [], _ -> true
            | _, _ -> false
        work root.Segments y.GetPathSegments 