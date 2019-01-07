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

An operations-centric is more flexible.

Base types `Document` and `Aggregate`.

`Document` *tracks* original file name and derived name after modifications.




