// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


module DocMake.WordBuilder

open Microsoft.Office.Interop

open DocMake.Base.Common
open DocMake.Base.OfficeUtils
open DocMake.Builder.BuildMonad
open DocMake.Builder.Document
open DocMake.Tasks



type WordRes = Word.Application

type WordBuild<'a> = BuildMonad<WordRes,'a>



let runWordBuild (env:Env) (ma:WordBuild<'a>) : 'a = 
    let wordApp = initWord ()
    let wordKill = fun (app:Word.Application) -> finalizeWord app
    consoleRun env wordApp wordKill ma


// *************************************
// Wraps DocMake.Tasks.DocFindReplace


/// DocFindReplace Api has more than one entry point...
let private docFindReplaceApi : DocFindReplace.DocFindReplaceApi<WordRes> = 
    DocFindReplace.makeAPI (fun (app:Word.Application) -> app)

let getTemplateDoc (docPath:string) : WordBuild<WordDoc> = 
    docFindReplaceApi.GetTemplateDoc docPath


let docFindReplace (searchList:SearchList) (template:WordDoc) : WordBuild<WordDoc> = 
    docFindReplaceApi.DocFindReplace searchList template


