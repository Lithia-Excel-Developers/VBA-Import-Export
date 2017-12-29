# VBA Import & Export Add-in for MS Excel

[![The MIT License](https://img.shields.io/badge/license-MIT-orange.svg?style=flat-square)](http://opensource.org/licenses/MIT)

The *VBA Import & Export Add-in* is an Add-in for *Microsoft Excel* that allows
you to import and export VBA code to and from Excel Workbooks (`.xlsm` files)
easily. This allows VBA code to be stored in plain text files alongside the
`.xlsm` file. This is essential for effectively utilizing Git in a VBA project
(or any other VCS).

## Installation

1. Download the Add-in: [VBA-Import-Export.xlam](https://github.com/mattpalermo/VBA-Import-Export/releases/download/v0.4.0/VBA-Import-Export.xlam) (version 0.4.0)
2. Add and Enable the Add-in in Excel
3. In Excel, Check the `Trust access to the VBA project model` check box
   located in `Trust Centre -> Trust Centre Settings -> Macro Settings ->
   Trust access to the VBA project model`.

## Usage

Menus for using the add-in can be found in the *Developer* tab of the *Excel
ribbon menu* and in the menu of the *VBA IDE*. Both menus provide the same
commands.

To get started quickly:
1. Save your `.xlsm` file in the directory where your VBA code will go.
2. Use the `Make Config File` command to make a `CodeExport.config.json` file in
that same directory. This records a list of VBA files and references.
3. Use the `Export` command to export the VBA code.
4. Notice the VBA code present in the same directory as your `.xlsm` file.
5. Save and close your Excel workbook.
6. (Optional) Commit the contents of your project directory into Git or any
   other VCS system.

When you return to work on the VBA project:
1. (Optional) Checkout a version of your project from Git or the VCS system you
   are using.
2. Open the Excel workbook (`.xlsm` file) in Excel.
3. Use the `Import` command to import the VBA code from the project directory\*
   and the references listed in the configuration file.
4. When you're ready to save: use the `Make Config File` command to update the
   config file; use the `Export` command to export the VBA code; then save
   and close the Excel workbook.

\* Only files listed in the configuration file will be imported.

## The configuration file

The `CodeExport.config.json` file declares what gets imported to and exported
from an Excel workbook (`.xlsm` file). The config file must be in the
same directory as the `.xlsm` file. The config file can be edited in
a text editor to make advanced adjustments that the `Make Config File` command
cannot do. A comprehensive example config file can be found at [test-projects/comprehensive/CodeExport.config.json](test-projects/comprehensive/CodeExport.config.json).

The config file uses the [JSON file format](https://en.wikipedia.org/wiki/JSON).
The following list describes the configuration properties that are used by
*VBA-Import-Export*:

* `VBAProject Name` - The name of the VBAProject. Will be set on import. Must
  not contain any spaces.
* `Module Paths` - A file system path for every VBA module that will be imported
  and exported by CodeExport. These may be relative or absolute paths.
* `Base Path` - A prefix to be prepended to all relative paths in
  `Module Paths`.
* `References` - A list of reference definitions. Each reference described will
  be referenced on import and dereferenced on export.

## Importing & Exporting

The `Import` command will:

* Import all the modules specified in the `Module Paths` configuration property.
Existing modules in the Excel file will be overwritten.
* Add all library references declared in the `References` configuration
property. Existing library references in the Excel file will be overwritten.
* Set the VBAProject name as declared in the `VBAProject Name` configuration
property.

The `Export` command will:

* Export all the modules specified in the `Module Paths` configuration property.
Existing files in the file system will be overwritten.
* Dereference libraries declared in the `References` configuration property.

## Support

You can submit questions, requests and bug reports to the
[issues list](https://github.com/mattpalermo/VBA-Import-Export/issues).
Github pull requests are also welcome.

## Authors and Attribution

* Scott Spence - Author, his version is at
[spences10/VBA-IDE-Code-Export](https://github.com/spences10/VBA-IDE-Code-Export)
* Matthew Palermo - Author, maintainer of this version
([mattpalermo/VBA-Import-Export](https://github.com/mattpalermo/VBA-Import-Export))
* Tim Hall - Author of the library [VBA-JSON](https://github.com/VBA-tools/VBA-JSON) which is used to read and write
the configuration file(s).
* All the authors of the clever VBA code snippets found in forums across the
internet which showed us how to solve some of the less documented nuances of VBA
and Excel. I wish I had kept a list.

## See Also

* [vba-blocks](https://www.vba-blocks.com/) - A VBA package manager in
development by Tim Hall. It will hopefully supersede this add-in.
* [VBA-IDE-Code-Export](https://github.com/spences10/VBA-IDE-Code-Export) - Scott
Spence's version of this add-in.
