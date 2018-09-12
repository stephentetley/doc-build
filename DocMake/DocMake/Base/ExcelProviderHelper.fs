// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

module DocMake.Base.ExcelProviderHelper



/// F# design guidelines say favour object-interfaces rather than records of functions...
type IExcelProviderHelper<'table,'row> = 
    abstract member ReadTableRows : 'table -> seq<'row>
    abstract member IsBlankRow: 'row -> bool

let excelReadRows (helper:IExcelProviderHelper<'table,'row>) (table:'table) : seq<'row> = 
    let allrows = helper.ReadTableRows table
    allrows |> Seq.filter (not << helper.IsBlankRow)


let excelReadRowsAsList (helper:IExcelProviderHelper<'table,'row>) (table:'table) : 'row list = 
    excelReadRows helper table |> Seq.toList

