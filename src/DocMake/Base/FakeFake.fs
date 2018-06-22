// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


/// Deprecate dependency on Fake as it has become too hard to
/// track its packages
module DocMake.Base.FakeFake

open System.IO


/// Note if the second path is prefixed by '\\'
/// "directory" </> "/file.ext" == "/file.ext"
let (</>) path1 path2 = Path.Combine(path1,path2)
