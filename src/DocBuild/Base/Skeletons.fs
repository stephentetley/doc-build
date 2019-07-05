// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Base

[<AutoOpen>]
module Skeletons = 

    open System.IO

    open DocBuild.Base


    let private getSourceSubdirectories () : DocMonad<string list, 'userRes> = 
        docMonad { 
            let! source = askSourceDirectory ()
            let! (kids: System.IO.DirectoryInfo[]) = 
                liftOperation "getSourceSubdirectories" <| fun _ -> System.IO.DirectoryInfo(source).GetDirectories() 
            return (kids |> Array.map (fun info -> info.Name) |> Array.toList)
        }


    type TestingSample = 
        | NoTestingProcesssAll
        | FilterDirectoryNames of filterCondition : (string -> bool)
        | TakeDirectories of count : int

    type SkeletonStepBeginMessage = int -> int -> string -> string

    type SkeletonStepFailMessage = string -> string

    type SkeletonOptions = 
        { CreateWorkingSubdirectory : bool 
          GenStepBeginMessage : SkeletonStepBeginMessage option
          GenStepFailMessage : SkeletonStepFailMessage option
          TestingSample : TestingSample
          ContinueOnFail : bool
        }

    let defaultSkeletonOptions:SkeletonOptions = 
        { CreateWorkingSubdirectory = true 
          GenStepBeginMessage = Some 
            <| fun ix count childFolderName -> 
                    sprintf "%i of %i: %s" ix count childFolderName
          GenStepFailMessage = Some <| sprintf "%s failed"
          TestingSample = NoTestingProcesssAll
          ContinueOnFail = true
        }

    /// Processing skeleton.
    /// For every child source folder (one level down) run the
    /// processing function on 'within' that folder. 
    /// With the option ``CreateWorkingSubdirectory = true`` 
    /// a subdirectory in be created in Working with the name 
    /// of the source (child) directory, and all output files 
    /// will be written there.
    let foreachSourceDirectory (skeletonOpts:SkeletonOptions) 
                               (process1: DocMonad<'a, 'userRes>) : DocMonad<unit, 'userRes> =  
        let strategy = 
            if skeletonOpts.CreateWorkingSubdirectory then 
                fun childDirectory action -> 
                    localSourceSubdirectory childDirectory 
                                            (localWorkingSubdirectory childDirectory action)
            else
                fun childDirectory action -> 
                    localSourceSubdirectory childDirectory action

        let sampleChildDirectories (children : string list) : string list= 
            match skeletonOpts.TestingSample with
            | NoTestingProcesssAll -> children
            | FilterDirectoryNames test -> 
                List.filter (fun path -> test (DirectoryInfo(path).Name)) children
            | TakeDirectories count -> List.take count children

        let logStepBegin (ix:int) (count:int) = 
            match skeletonOpts.GenStepBeginMessage with
            | None -> mreturn ()
            | Some genMessage -> 
                docMonad { 
                   let! kid = askSourceDirectory () |>> Internal.FilePaths.getPathName1
                   let message = genMessage ix count kid 
                   do (printfn "%s" message)
                   do! tellLine message
                   return ()
                }
        let logStepFail () = 
            match skeletonOpts.GenStepFailMessage with
            | None -> mreturn ()
            | Some genMessage -> 
                docMonad { 
                    let! kid = askSourceDirectory () |>> fileObjectName
                    let message = genMessage kid
                    do (printfn "%s" message)
                    do! tellLine message
                    return ()
                }
        let proceedM (proc:DocMonad<unit, 'userRes>) : DocMonad<unit, 'userRes> = 
            docMonad { 
                match! (captureM proc) with
                | Error msg -> 
                    if skeletonOpts.ContinueOnFail then 
                        do! logStepFail () 
                        return ()
                    else 
                        logStepFail () .>> docError "Build step failed" |> ignore
                | Ok _ -> return ()
                }
        let processChildDirectory (ix:int) (count:int) : DocMonad<unit, 'userRes> = 
            docMonad { 
                do! logStepBegin ix count
                return! proceedM (ignoreM process1)
            }

        getSourceSubdirectories () >>= fun srcDirs -> 
        let sources = sampleChildDirectories srcDirs
        let count = List.length sources
        foriMz sources 
               (fun ix dir -> strategy dir (processChildDirectory (ix + 1) count))

