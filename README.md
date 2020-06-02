# ExcelDna-Unpack

ExcelDna-Unpack is a command-line utility to extract the contents of Excel-DNA add-ins that have been packed with ExcelDnaPack

## Give a Star! :star:

If you like or are using this project please give it a star. Thanks!

## Getting Started

Download the latest version of ExcelDna-Unpack from the [Releases tab](https://github.com/augustoproiete/exceldna-unpack/releases), and extract the file `exceldna-unpack-{version}{rid}.zip` to the location where your `.xll` add-in is stored.

Run `exceldna-unpack.exe --xllFile=NameOfYourAddIn.xll`

```
Usage: exceldna-unpack.exe [<options>]

Where [<options>] is any of:

--xllFile=VALUE    The XLL file to be unpacked; e.g. MyAddIn-packed.xll
--outFolder=VALUE  [Optional] The folder into which the extracted files will be written; defaults to '.\unpacked'
--overwrite        [Optional] Allow existing files of the same name to be overwritten

Example: exceldna-unpack.exe --xllFile=MyAddins\FirstAddin-packed.xll
         The extracted files will be saved to MyAddins\unpacked
```

## Release History

Click on the [Releases](https://github.com/augustoproiete/exceldna-unpack/releases) tab on GitHub.


---

_Copyright &copy; 2014-2020 C. Augusto Proiete & Contributors - Provided under the [Apache License, Version 2.0](LICENSE)._
