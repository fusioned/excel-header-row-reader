# Excel header parser

Quick Python 3 script to read the first row from each workbook and worksheet 
(and optionally, the next line) in CSV format.

The filename and sheet name are included in the output.

## Usage

I combine this with shell commands 

e.g. to send all the workbooks in old to new date order:

```
ls -1trQ *.xlsx |  xargs  ./read_excel_headers.py
```

The filenames are quoted (-Q), one per line (-1), concatenated into a single line (xargs) and parsed by the script
