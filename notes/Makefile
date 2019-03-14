
all: notes

notes:
	pandoc --from=markdown --pdf-engine=pdflatex --standalone --output=output/notes.pdf notes.md
	pandoc --from=markdown+pandoc_title_block --to=html --metadata pagetitle="Notes" --standalone --output=output/notes.html notes.md
	pandoc --from=markdown --to=docx --reference-doc=include/custom-reference1.docx --standalone --output=output/notes.docx notes.md

