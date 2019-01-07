@echo on

REM notes
pandoc notes.md -s -o notes.pdf

pandoc notes.md --metadata pagetitle="Notes" -f markdown+pandoc_title_block -t html -s -o notes.html

