// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Base

[<AutoOpen>]
module Skeletons = 

    open DocBuild.Base


    let private getSourceSubdirectories () : DocMonad<'userRes,string list> = 
        docMonad { 
            let! source = askSourceDirectory ()
            let! (kids: System.IO.DirectoryInfo[]) = 
                liftOperation "getSourceSubdirectories" <| fun _ -> System.IO.DirectoryInfo(source).GetDirectories() 
            return (kids |> Array.map (fun info -> info.Name) |> Array.toList)
        }


    type SkeletonStepBeginMessage = int -> int -> string -> string

    type SkeletonStepFailMessage = string -> string

    type SkeletonOptions = 
        { GenStepBeginMessage: SkeletonStepBeginMessage option
          GenStepFailMessage: SkeletonStepFailMessage option
          DebugSelectSample: (string list -> string list) option
          ContinueOnFail: bool
        }

    let defaultSkeletonOptions:SkeletonOptions = 
        { GenStepBeginMessage = Some 
            <| fun ix count childFolderName -> 
                    sprintf "%i of %i: %s" ix count childFolderName
          GenStepFailMessage = Some <| sprintf "%s failed"
          DebugSelectSample = None
          ContinueOnFail = true
        }

    let private runSkeleton (skeletonOpts:SkeletonOptions) 
                            (strategy: string -> DocMonad<'userRes,unit> -> DocMonad<'userRes,unit>)
                            (process1: DocMonad<'userRes,'a>) : DocMonad<'userRes, unit> =  
        let processZ: DocMonad<'userRes, unit> = process1 |>> fun _ -> ()
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
        let proceedM (proc:DocMonad<'userRes,unit>) : DocMonad<'userRes, unit> = 
            docMonad { 
                match! (optionMaybeM proc) with
                | None -> 
                    if skeletonOpts.ContinueOnFail then 
                        return ()
                    else 
                        logStepFail () .>> docError "Build step failed" |> ignore
                | Some _ -> return ()
                }
        let processChildDirectory (ix:int) (count:int) : DocMonad<'userRes, unit> = 
            docMonad { 
                do! logStepBegin ix count
                return! (proceedM processZ)
            }

        getSourceSubdirectories () >>= fun srcDirs -> 
        let sources = filterChildDirectories srcDirs
        let count = List.length sources
        foriMz sources 
               (fun ix dir -> strategy dir (processChildDirectory (ix + 1) count))

    
    /// Processing skeleton.
    /// For every child source folder (one level down) run the
    /// processing function on 'within' that folder. 
    /// Generate the results in a child folder of the same name under
    /// the working folder.
    let foreachSourceIndividualOutput (skeletonOpts:SkeletonOptions) 
                                      (process1: DocMonad<'userRes,'a>) : DocMonad<'userRes, unit> = 
        let strategy = fun childDirectory action -> 
                localSourceSubdirectory childDirectory 
                                        (localWorkingSubdirectory childDirectory action)
        runSkeleton skeletonOpts strategy process1

    /// Processing skeleton.
    /// For every child source folder (one level down) run the
    /// processing function on 'within' that folder. 
    /// Generate the results in the top level working folder.
    let foreachSourceCommonOutput (skeletonOpts:SkeletonOptions) 
                                  (process1: DocMonad<'userRes,'a>) : DocMonad<'userRes, unit> = 
        let strategy = fun childDirectory action -> 
                localSourceSubdirectory childDirectory action
        runSkeleton skeletonOpts strategy process1