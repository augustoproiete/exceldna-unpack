# ExcelDnaUnpack

ExcelDnaUnpack is a command-line utility to extract the contents of ExcelDna add-ins packed with ExcelDnaPack

```
Usage: ExcelDnaUnpack.exe [<options>]

Where [<options>] is any of:

--xllFile=VALUE    The XLL file to be unpacked; e.g. MyAddIn-packed.xll
--outFolder=VALUE  [Optional] The folder into which the extracted files will be written; defaults to '.\unpacked'
--overwrite        [Optional] Allow existing files of the same name to be overwritten

Example: ExcelDnaUnpack.exe --xllFile=MyAddins\FirstAddin-packed.xll
         The extracted files will be saved to MyAddins\unpacked
```

## Give a Star! :star:

If you like or are using this project please give it a star. Thanks!

## Release History

Click on the [Releases](https://github.com/augustoproiete/ExcelDnaUnpack/releases) tab on GitHub.


---

_Copyright &copy; 2014-2020 C. Augusto Proiete & Contributors - Provided under the [Apache License, Version 2.0](LICENSE)._
