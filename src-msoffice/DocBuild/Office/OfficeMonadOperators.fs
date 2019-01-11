// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Office

// Explicitly import this module:
// open DocBuild.Office.OfficeMonadOperators

module OfficeMonadOperators = 


    open DocBuild.Office.OfficeMonad

    // ****************************************************
    // Errors

    /// Operator for swapError
    let (<?&>) (msg:string) (ma:OfficeMonad<'a>) : OfficeMonad<'a> = 
        swapError msg ma

    /// Operator for flip swapError
    let (<&?>) (ma:OfficeMonad<'a>) (msg:string) : OfficeMonad<'a> = 
        swapError msg ma


    // ****************************************************
    // Monadic operations

    /// Bind operator
    let (>>=) (ma:OfficeMonad<'a>) (fn:'a -> OfficeMonad<'b>) : OfficeMonad<'b> = 
        officeMonad.Bind(ma,fn)

    /// Flipped Bind operator
    let (=<<) (fn:'a -> OfficeMonad<'b>) (ma:OfficeMonad<'a>) : OfficeMonad<'b> = 
        officeMonad.Bind(ma,fn)


    /// Operator for fmap.
    let (|>>) (ma:OfficeMonad<'a>) (fn:'a -> 'b) : OfficeMonad<'b> = 
        fmapM fn ma

    /// Flipped fmap.
    let (<<|) (fn:'a -> 'b) (ma:OfficeMonad<'a>) : OfficeMonad<'b> = 
        fmapM fn ma

    /// Operator for altM
    let (<||>) (ma:OfficeMonad<'a>) (mb:OfficeMonad<'a>) : OfficeMonad<'a> = 
        altM ma mb 


    /// Operator for apM
    let (<**>) (ma:OfficeMonad<'a -> 'b>) (mb:OfficeMonad<'a>) : OfficeMonad<'b> = 
        apM ma mb

    /// Operator for fmapM
    let (<&&>) (fn:'a -> 'b) (ma:OfficeMonad<'a>) : OfficeMonad<'b> = 
        fmapM fn ma



    /// Operator for seqL
    let (.>>) (ma:OfficeMonad<'a>) (mb:OfficeMonad<'b>) : OfficeMonad<'a> = 
        seqL ma mb

    /// Operator for seqR
    let (>>.) (ma:OfficeMonad<'a>) (mb:OfficeMonad<'b>) : OfficeMonad<'b> = 
        seqR ma mb



    /// Operator for kleisliL
    let (>=>) (mf : 'a -> OfficeMonad<'b>)
              (mg : 'b -> OfficeMonad<'c>)
              (source:'a) : OfficeMonad<'c> = 
        kleisliL mf mg source


    /// Operator for kleisliR
    let (<=<) (mf : 'b -> OfficeMonad<'c>)
              (mg : 'a -> OfficeMonad<'b>)
              (source:'a) : OfficeMonad<'c> = 
        kleisliR mf mg source




