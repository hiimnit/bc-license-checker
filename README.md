# bc-license-checker

Simple command line util to compare detailed permission report for a Business Central/Dynamics NAV license
with list of objects xlsx format.

```
Usage: bc-license-checker --license <LICENSE> --objects <OBJECTS>

Options:
  -l, --license <LICENSE>  Path to detailed permission report text file
  -o, --objects <OBJECTS>  Path to exported objects in xlsx format
  -h, --help               Print help information
  -V, --version            Print version information
```

## Expected permission report format

Expected input is the detailed permission report generated in ms business center.

## Expected xlsx format

Object type | Object id | Object name
------------|-----------|------------
TableData   | 50000     | Name
Page        | 50000     | Name
Report      | 50000     | Name
Codeunit    | 50000     | Name
XMLPort     | 50000     | Name
