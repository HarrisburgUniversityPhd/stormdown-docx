# stormdown

A Harrisburg University Dissertation Template.
This version is the Microsoft Word version.
Legacy versions exist for [HTML version](https://github.com/HarrisburgUniversityPhd/stormdown-html) and [LaTeX](https://github.com/HarrisburgUniversityPhd/stormdown-latex).
Please note that those versions are _legacy_.
It is unlikely that future updates will be made to them.

# Requirements

For Users

* Microsoft Word

For Developers

* [VS Code](https://code.visualstudio.com/)
* [7-Zip](https://www.7-zip.org/)
* [Computer Modern](https://www.fontsquirrel.com/fonts/computer-modern) font 

# How to update the template

_.dotx_ files are _pkzip_ files with a different extension.
_pkzip_ files are binary, so they don't work well with Git.
In order to work on the templates do the following:

1. Use 7-Zip to convert the known folder + _.xml_ structure to a _.dotx_ file.
2. Use MS Word to update the template.
   **NOTE**: Open with drag and drop.
   If you open using double-click, MS Word assumes you want to make a document _based_ on the template, not update the template itself.
3. Use 7-Zip to convert the _.dotx_ back to a folder + _.xml_ structure.
4. Pretty print the _.xml_.
4. Use VS Code to update source control.
5. Copy the _.dotx_ file to [releases](https://github.com/HarrisburgUniversityPhd/stormdown-docx/releases).

```{powershell}
cd "C:/repos/HarrisburgUniversityPhd/stormdown-docx"
$7z = 'C:/Program Files/7-Zip/7z.exe'
$name = 'base'

. $7z a -tzip -mx1 -r "templates/$($name).dotx" "./templates/$name/*.*"
# edit using Word
rm "templates/$name" -r -force
. $7z x -y $("-otemplates/$name") "templates/$($name).dotx"
dir -r "templates/$($name)/*.xml" | % { ([xml](gc -LiteralPath $_ -encoding utf8)).Save($_) }
dir -r "templates/$($name)/*.rels" | % { ([xml](gc -LiteralPath $_ -encoding utf8)).Save($_) }
```
