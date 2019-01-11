// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Base

// Explicitly import this module:
// open DocBuild.Base.DocMonadOperators

module DocMonadOperators = 


    open DocBuild.Base
    open DocBuild.Base.DocMonad

    // ****************************************************
    // Errors

    /// Operator for swapError
    let (<?&>) (msg:string) (ma:DocMonad<'a>) : DocMonad<'a> = 
        swapError msg ma

    /// Operator for flip swapError
    let (<&?>) (ma:DocMonad<'a>) (msg:string) : DocMonad<'a> = 
        swapError msg ma


    // ****************************************************
    // Monadic operations

    /// Bind operator
    let (>>=) (ma:DocMonad<'a>) (fn:'a -> DocMonad<'b>) : DocMonad<'b> = 
        docMonad.Bind(ma,fn)

    /// Flipped Bind operator
    let (=<<) (fn:'a -> DocMonad<'b>) (ma:DocMonad<'a>) : DocMonad<'b> = 
        docMonad.Bind(ma,fn)


    /// Operator for fmap.
    let (|>>) (ma:DocMonad<'a>) (fn:'a -> 'b) : DocMonad<'b> = 
        fmapM fn ma

    /// Flipped fmap.
    let (<<|) (fn:'a -> 'b) (ma:DocMonad<'a>) : DocMonad<'b> = 
        fmapM fn ma

    /// Operator for altM
    let (<||>) (ma:DocMonad<'a>) (mb:DocMonad<'a>) : DocMonad<'a> = 
        altM ma mb 


    /// Operator for apM
    let (<**>) (ma:DocMonad<'a -> 'b>) (mb:DocMonad<'a>) : DocMonad<'b> = 
        apM ma mb

    /// Operator for fmapM
    let (<&&>) (fn:'a -> 'b) (ma:DocMonad<'a>) : DocMonad<'b> = 
        fmapM fn ma



    /// Operator for seqL
    let (.>>) (ma:DocMonad<'a>) (mb:DocMonad<'b>) : DocMonad<'a> = 
        seqL ma mb

    /// Operator for seqR
    let (>>.) (ma:DocMonad<'a>) (mb:DocMonad<'b>) : DocMonad<'b> = 
        seqR ma mb



    /// Operator for kleisliL
    let (>=>) (mf : 'a -> DocMonad<'b>)
              (mg : 'b -> DocMonad<'c>)
              (source:'a) : DocMonad<'c> = 
        kleisliL mf mg source


    /// Operator for kleisliR
    let (<=<) (mf : 'b -> DocMonad<'c>)
              (mg : 'a -> DocMonad<'b>)
              (source:'a) : DocMonad<'c> = 
        kleisliR mf mg source




