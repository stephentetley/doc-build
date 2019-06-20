// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Base

[<AutoOpen>]
module Skeletons = 

    open DocBuild.Base


    let private getSourceSubdirectories () : DocMonad<string list, 'userRes> = 
        docMonad { 
            let! source = askSourceDirectory ()
            let! (kids: System.IO.DirectoryInfo[]) = 
                liftOperation "getSourceSubdirectories" <| fun _ -> System.IO.DirectoryInfo(source).GetDirectories() 
            return (kids |> Array.map (fun info -> info.Name) |> Array.toList)
        }


    type SkeletonStepBeginMessage = int -> int -> string -> string

    type SkeletonStepFailMessage = string -> string

    type SkeletonOptions = 
        { CreateWorkingSubdirectory : bool 
          GenStepBeginMessage : SkeletonStepBeginMessage option
          GenStepFailMessage : SkeletonStepFailMessage option
          DebugSelectSample : (string list -> string list) option
          ContinueOnFail : bool
        }

    let defaultSkeletonOptions:SkeletonOptions = 
        { CreateWorkingSubdirectory = true 
          GenStepBeginMessage = Some 
            <| fun ix count childFolderName -> 
                    sprintf "%i of %i: %s" ix count childFolderName
          GenStepFailMessage = Some <| sprintf "%s failed"
          DebugSelectSample = None
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
        let ignoreM : DocMonad<unit, 'userRes> = process1 |>> fun _ -> ()

        let strategy = 
            if skeletonOpts.CreateWorkingSubdirectory then 
                fun childDirectory action -> 
                    localSourceSubdirectory childDirectory 
                                            (localWorkingSubdirectory childDirectory action)
            else
                fun childDirectory action -> 
                    localSourceSubdirectory childDirectory action

        let filterChildDirectories = 
            match skeletonOpts.DebugSelectSample with
            | None -> id
            | Some fn -> fn
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
                match! (optionMaybeM proc) with
                | None -> 
                    if skeletonOpts.ContinueOnFail then 
                        return ()
                    else 
                        logStepFail () .>> docError "Build step failed" |> ignore
                | Some _ -> return ()
                }
        let processChildDirectory (ix:int) (count:int) : DocMonad<unit, 'userRes> = 
            docMonad { 
                do! logStepBegin ix count
                return! (proceedM ignoreM)
            }

        getSourceSubdirectories () >>= fun srcDirs -> 
        let sources = filterChildDirectories srcDirs
        let count = List.length sources
        foriMz sources 
               (fun ix dir -> strategy dir (processChildDirectory (ix + 1) count))

