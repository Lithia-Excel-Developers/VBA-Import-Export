# VBA Import & Export Add-in for MS Excel

The *VBA Import & Export Add-in* is an Add-in for *Microsoft Excel* that allows
you to import and export VBA code to and from Excel Workbooks (`.xlsm` files)
easily. This allows VBA code to be stored in plain text files alongside the
`.xlsm` file. This is essential for effectively utilizing Git in a VBA project
(or any other VCS).

## Installation

1. Download the Add-in: [VBA-Import-Export.xlam](https://github.com/mattpalermo/VBA-Import-Export/releases/download/v0.4.0/VBA-Import-Export.xlam) (version 0.4.0)
2. Add and enable the Add-in in Excel
3. In Excel, Check the `Trust access to the VBA project model` check box
   located in `Trust Centre -> Trust Centre Settings -> Macro Settings ->
   Trust access to the VBA project model`.

## Usage

The add-in can be used from *Developer* tab of the *Excel
ribbon menu* and in the menu of the *VBA IDE*. Both menus provide the same
commands.

### Getting started
1. Save your `.xlsm` file into a folder.
2. Use the `Make Config File` command to make a `CodeExport.config.json` file in
the same folder as the `.xlsm` file. This records a list of VBA files and
references.
3. Use the `Export` command to export the VBA code. Notice the VBA code present
in the same folder as your `.xlsm` file.
4. Save and close your Excel workbook.
5. *(Optional)* Commit the contents of your project directory into Git or any
   other VCS system.

**Important:**
[VBA-project-template](https://github.com/mattpalermo/VBA-project-template)
provides config files to ensure Git and text editors play nicely with your
project files.

When you return to work on the VBA project:
1. *(Optional)* Checkout a version of your project from Git or the VCS system you
   are using.
2. Open the Excel workbook (`.xlsm` file) in Excel.
3. Use the `Import` command to import the VBA code from the project directory\*
   and the references listed in the configuration file.
4. Regularly use the `Save` command to save your changes to the file system while you work.
5. Use the `Make Config File` command to update the config file when modules or references are added or removed.
4. When you're finished, use the `Export` command to export your work, then save
   and close the Excel workbook.

\* Only files listed in the configuration file will be imported.

### Safety tips

Here's some tips to avoid loosing data while using this Add-in:

* Do regular backups! Use the method that you won't forget. My favourite method is to use Git. Any versioning system would also work.
* If you make changes in the Excel document, don't edit the files in the file system before you `Export`. The inevitable `Export` will overwrite your changes.
* If you make changes in the Excel document, don't `Import` before using `Save` or `Export`. You will just overwrite your changes with what you started with.
* `Save` regularly to avoid making the mistake above.

## The configuration file

The `CodeExport.config.json` file declares what gets imported and exported from
an Excel workbook. The config file must be in the same directory as the `.xlsm`
file. The config file can be edited in a text editor to make advanced
adjustments that the `Make Config File` command cannot do. A comprehensive
example config file can be found at
[test-projects/comprehensive/CodeExport.config.json](test-projects/comprehensive/CodeExport.config.json).
The config file uses the [JSON file format](https://en.wikipedia.org/wiki/JSON)
and the configuration properties are:

* `VBAProject Name` - The name of the VBAProject. Will be set on import. Must
  not contain any spaces.
* `Module Paths` - A file system path for every VBA module that will be imported
  and exported by CodeExport. These may be relative or absolute paths.
* `Base Path` - A prefix to be prepended to all relative paths in
  `Module Paths`.
* `References` - A list of reference definitions. Each reference described will
  be referenced on import and dereferenced on export.

## Importing, Saving & Exporting

The `Import` command will:

* Import all the modules specified in the `Module Paths` configuration property.
Existing modules in the Excel file will be overwritten.
* Add all library references declared in the `References` configuration
property. Existing library references in the Excel file will be overwritten.
* Set the VBAProject name as declared in the `VBAProject Name` configuration
property.

The `Save` command will:

* Export all the modules specified in the `Module Paths` configuration property.
Existing files in the file system will be overwritten.

The `Export` command will:

* Do the same as the `Save` command.
* Dereference libraries declared in the `References` configuration property.

## Support

You can submit questions, requests and bug reports to the
[issues list](https://github.com/mattpalermo/VBA-Import-Export/issues).
Github pull requests are also welcome.

## Authors and Attribution

* Scott Spence - Author
([spences10/VBA-IDE-Code-Export](https://github.com/spences10/VBA-IDE-Code-Export))
* Matthew Palermo - Author
([mattpalermo/VBA-Import-Export](https://github.com/mattpalermo/VBA-Import-Export))
* Tim Hall - Author of the library [VBA-JSON](https://github.com/VBA-tools/VBA-JSON)
* [Kevin Conner](https://github.com/connerk) - Author of the Save action

## See Also

* [vba-blocks](https://www.vba-blocks.com/) - A VBA package manager in
development by Tim Hall. It will hopefully supersede this add-in.
* [VBA-IDE-Code-Export](https://github.com/spences10/VBA-IDE-Code-Export) - Scott
Spence's version of this add-in.
