# Notes

An "operations-centric" API is nicer than an "options-centric" one.

Operations-centric:

~~~~~
PhotoCollection "path/to/photos" |> optimize |> photoPdf "MyPhotos.pdf"
~~~~~

Options-centric: 

~~~~~
PhotoCollection(Path = "path/to/photos", Optimize=true, Output= "MyPhotos.pdf")
~~~~~

An operations-centric API is more flexible.

Base types `Document` and `Aggregate`.

A `Document` tracks the original file name and the derived temporary name after modifications. No modifications are performed on the original file.




