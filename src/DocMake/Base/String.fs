// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause

module DocMake.Base.String


let isPrefixOf (source:string) (needle:string) : bool = 
    let len = needle.Length
    needle.Equals (source.Substring(0,len))
    

let isSuffixOf (source:string) (needle:string) : bool = 
    let lenS = source.Length
    let lenT = needle.Length
    needle.Equals (source.Substring(lenS - lenT))


let isPrefixOfBy (source:string) (needle:string) (comparisonType:System.StringComparison) : bool = 
    let len = needle.Length
    needle.Equals (source.Substring(0,len), comparisonType)


let isSuffixOfBy (source:string) (needle:string) (comparisonType:System.StringComparison) : bool = 
    let lenS = source.Length
    let lenT = needle.Length
    needle.Equals (source.Substring(lenS - lenT), comparisonType) 


let leftOf (source:string) (needle:string) : string = 
    try
        let ix = source.IndexOf(needle) 
        source.Substring(0,ix)
    with
    | :? System.ArgumentOutOfRangeException -> ""

let rightOf (source:string) (needle:string) : string = 
    try
        let ix = source.IndexOf(needle) + needle.Length
        source.Substring(ix)
    with
    | :? System.ArgumentOutOfRangeException -> ""

let leftOfBy (source:string) (needle:string) (comparisonType:System.StringComparison) : string = 
    try
        let ix = source.IndexOf(value = needle, comparisonType = comparisonType) 
        source.Substring(0,ix)
    with
    | :? System.ArgumentOutOfRangeException -> ""


let rightOfBy (source:string) (needle:string) (comparisonType:System.StringComparison) : string = 
    try
        let ix = source.IndexOf(value = needle, comparisonType = comparisonType) + needle.Length
        source.Substring(ix)
    with
    | :? System.ArgumentOutOfRangeException -> ""

let between (source:string) (needleLeft:string) (needleRight:string) : string = 
    leftOf (rightOf source needleLeft) needleRight

let betweenBy (source:string) (needleLeft:string) (needleRight:string) (comparisonType:System.StringComparison) : string = 
    leftOfBy (rightOfBy source needleLeft comparisonType) needleRight comparisonType
