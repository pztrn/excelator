# Excelator

Simple program in Go that will filter data from Excel file based on passed
keyword.

It will take a good look thru whole xlsx file and produce file with
filtered data.

# Usage

Program takes two parameters:

* ``-filename`` - path to file. Remember to quote it if it contains spaces!
* ``-keyword`` - keyword to search for.

Excelator will iterate thru all filled lines and cells and search for a
keyword. If it found a match - it will write this line to file
``filtered.xlsx`` which will be placed near binary (or in ``$PWD``).
