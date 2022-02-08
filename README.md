# TranslationEditor

This repository contains the VBA code for the XLSM VVVVVV translation editor. The full versions of the XLSM can be found under Releases.

The .bas files are in the repo itself both for interest, and to provide diffs/comparisons between different versions, but they're not that useful by themselves unless you recreate the worksheets (it's a "Controls" sheet with buttons and a few cells with a special purpose, and then a separate sheet named after each `.xml` file in a language pack). All I did to get the VBA out of the Excel file was export each module from Excel itself (Ctrl+E).


# Changelog

2022-02-07
- Accommodate `expect` field in `strings_plural.xml`

2022-01-22
- Initial version
