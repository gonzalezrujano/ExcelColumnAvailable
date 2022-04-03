# Calculator Excel columns keys availables

## Example Excel Sheet
<pre>
-------------------------------------------------------
  | A               | B              | C              |
-------------------------------------------------------
1 | Column1         |                |                |
-------------------------------------------------------
2 | Value           | Value          | Value          |
-------------------------------------------------------
3 | ...             | ...            | ...            |
-------------------------------------------------------
</pre>

## Input
```
new_column_names = ["Column2", "Column3"]
last_column_usage: = "A"
keys_column_availables = ExcelColumnAvailables(new_column_names, last_column_usage)
```

## Output
```
keys_column_availables.get_columns()
[
    {'column_name': 'Column2', 'column_key': 'B'},
    {'column_name': 'Column3', 'column_key': 'C'}
]
```
