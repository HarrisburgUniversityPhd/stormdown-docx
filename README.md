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
We need the XML contents of the _pkzip_ file expanded to a folder structure to work well with Git.

In order to work on the templates do the following:

1. Create a fresh _.dotx_ file from folder structure.
   (see below).
2. Use MS Word to update the template.
   **NOTE**: Open the file with drag and drop OR right-click > open.
   Do NOT use double-click.
   If you open using double-click, MS Word assumes you want to make a document _based_ on the template, not update the template itself.
3. Convert the single _.dotx_ back to the folder structure.
   (see below).
4. Use VS Code to update source control.
5. Copy the _.dotx_ file to [releases](https://github.com/HarrisburgUniversityPhd/stormdown-docx/releases).

```{powershell}
cd "C:/repos/HarrisburgUniversityPhd/stormdown-docx/templates"
$policy = get-ExecutionPolicy –Scope Process
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned –Scope Process -Force
. ../tools/tools.ps1
Set-ExecutionPolicy -ExecutionPolicy $policy –Scope Process -Force

write-template -name 'base'
# edit using Word
write-folder -name 'base'
```
