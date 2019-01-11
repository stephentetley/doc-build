// Copyright (c) Stephen Tetley 2019
// License: BSD 3 Clause


namespace DocBuild.Base

// Explicitly import this module:
// open DocBuild.Base.DocMonadOperators

module DocMonadOperators = 


    open DocBuild.Base.DocMonad

    // ****************************************************
    // Errors

    /// Operator for swapError
    let (<?&>) (msg:string) (ma:DocMonad<'res,'a>) : DocMonad<'res,'a> = 
        swapError msg ma

    /// Operator for flip swapError
    let (<&?>) (ma:DocMonad<'res,'a>) (msg:string) : DocMonad<'res,'a> = 
        swapError msg ma


    // ****************************************************
    // Monadic operations

    /// Bind operator
    let (>>=) (ma:DocMonad<'res,'a>) 
              (fn:'a -> DocMonad<'res,'b>) : DocMonad<'res,'b> = 
        docMonad.Bind(ma,fn)

    /// Flipped Bind operator
    let (=<<) (fn:'a -> DocMonad<'res,'b>) 
              (ma:DocMonad<'res,'a>) : DocMonad<'res,'b> = 
        docMonad.Bind(ma,fn)


    /// Operator for fmap.
    let (|>>) (ma:DocMonad<'res,'a>) (fn:'a -> 'b) : DocMonad<'res,'b> = 
        fmapM fn ma

    /// Flipped fmap.
    let (<<|) (fn:'a -> 'b) (ma:DocMonad<'res,'a>) : DocMonad<'res,'b> = 
        fmapM fn ma

    /// Operator for altM
    let (<||>) (ma:DocMonad<'res,'a>) 
               (mb:DocMonad<'res,'a>) : DocMonad<'res,'a> = 
        altM ma mb 


    /// Operator for apM
    let (<**>) (ma:DocMonad<'res,'a -> 'b>) 
               (mb:DocMonad<'res,'a>) : DocMonad<'res,'b> = 
        apM ma mb

    /// Operator for fmapM
    let (<&&>) (fn:'a -> 'b) (ma:DocMonad<'res,'a>) : DocMonad<'res,'b> = 
        fmapM fn ma



    /// Operator for seqL
    let (.>>) (ma:DocMonad<'res,'a>) 
              (mb:DocMonad<'res,'b>) : DocMonad<'res,'a> = 
        seqL ma mb

    /// Operator for seqR
    let (>>.) (ma:DocMonad<'res,'a>) 
              (mb:DocMonad<'res,'b>) : DocMonad<'res,'b> = 
        seqR ma mb



    /// Operator for kleisliL
    let (>=>) (mf : 'a -> DocMonad<'res,'b>)
              (mg : 'b -> DocMonad<'res,'c>)
              (source:'a) : DocMonad<'res,'c> = 
        kleisliL mf mg source


    /// Operator for kleisliR
    let (<=<) (mf : 'b -> DocMonad<'res,'c>)
              (mg : 'a -> DocMonad<'res,'b>)
              (source:'a) : DocMonad<'res,'c> = 
        kleisliR mf mg source




