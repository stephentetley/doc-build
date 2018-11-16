// Copyright (c) Stephen Tetley 2018
// License: BSD 3 Clause


module DocBuild.JpegDoc



open DocBuild.Internal.ImageMagickUtils



// NOTEs 
// JpegDoc exists to support manipulation by ImageMagick
// (usually auto-orientation, resizing...).
// ImageMagick works destructively on an image handle, we can
// save a new copy after applying each trafo but that generates 
// a lot of junk.
//
// In ideal API might look something like:
//
// (new JpegDoc("DSC01.jpg")).AutoOrient().Resize(640,480).SaveAs("new.jpg")
//
// But these operations ``AutoOrient().Resize(640,480).SaveAs("new.jpg")``
// Clearly work on a "handle" not on the file.
//
// (new JpegDoc("DSC01.jpg")).Transform(fun img -> img.AutoOrient().Resize(640,480).SaveAs("new.jpg"))
//
// One significant problem is that ImageMagick extracts image info from the 
// file, not the image handle.
//
// Maybe to answer is to destructively operate on a single temp file



type JpegDoc = 
    val private JpegPath : string

    new (filePath:string) = 
        { JpegPath = filePath }

    member internal v.Body 
        with get() : string = v.JpegPath



let jpegDoc (path:string) : JpegDoc = new JpegDoc (filePath = path)

