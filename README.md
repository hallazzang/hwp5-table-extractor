# hwp5-table-extractor

hwp5-table-extractor is a tool for extracting tables from Hwp5 file.
It is developed in Python 3.6.1.

## Screenshot

![screenshot](screenshot.png)

Left: Rendered HTML
Right: Hwp Viewer for Mac.

## Dependencies

- olefile
- click
- jinja2

## Usage

Currently, no installation script is provided.
Just clone this repository and install dependencies, then run it manually:

```bash
$ git clone https://github.com/hallazzang/hwp5-table-extractor.git
$ cd hwp5-table-extractor
$ virtualenv -p python3 venv
$ source venv/bin/activate
(venv) $ pip install -r requirements.txt
(venv) $ python cli.py <INPUT_FILE> <OUTPUT_FILE>
```

## Notes

Supported output format is HTML only for now, but you can still access to the table
structure through `Table` object. It has `list` of rows and each row consists
of `list` of `TableCell`s.

So, the entire structure looks like:

```
<class Table>
    .row_cnt = XX
    .col_cnt = XX
    .rows = [
        [<class TableCell>, <class TableCell>, ...],
        [<class TableCell>, <class TableCell>, ...],
        ...
    ]

<class TableCell>
    .lines = ['Line 1', 'Line 2', 'Line 3', ...]
    .row = XX
    .col = XX
    .row_span = XX
    .col_span = XX
```

Note that each row can have different count of cell because of the `row_span`s
and `col_span`s.