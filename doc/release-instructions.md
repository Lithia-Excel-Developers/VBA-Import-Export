# Release Instructions

Steps to take to create a new release:

1. Choose a version number following the
   [semantic versioning](http://semver.org) guidelines.
2. Update [CHANGELOG.md](../CHANGELOG.md).
3. Build the `VBA-Import-Export.xlsm` workbook. Don't forget to set a
   VBA Project password of "123".
4. From this, create the `VBA-Import-Export.xlam` file.
5. Install `VBA-Import-Export.xlam` and test on
   [test-projects/comprehensive](../test-projects/comprehensive)
6. Create a new GitHub release:
    * Create a new tag following a format similar to `v1.2.3` (substituting the
      appropriate version number).
    * Title the release following a format similar to `Version 1.2.3`
      (substituting the appropriate version number).
    * Attach `VBA-Import-Export.xlam` as a downloadable binary.
7. Cheer and celebrate loudly about another great release. Let it be known!
