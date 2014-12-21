ExcelDnaUnpack
==============

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

## Release History

 * **[v1.0.0](https://github.com/CaioProiete/ExcelDnaUnpack/releases/tag/v1.0.0)** - 2014-12-21
   - Initial version

## License   
Copyright 2014 Caio Proiete

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

[http://www.apache.org/licenses/LICENSE-2.0](http://www.apache.org/licenses/LICENSE-2.0)

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
