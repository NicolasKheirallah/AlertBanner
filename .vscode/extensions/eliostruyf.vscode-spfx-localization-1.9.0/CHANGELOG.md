# Change Log

# [1.9.0] - 2022-01-09

- [#26](https://github.com/estruyf/vscode-spfx-localization/issues/26): Excel format support. Thanks to [Nikolay Belykh](https://github.com/nbelyh).

# [1.8.0] - 2022-12-22

- [#24](https://github.com/estruyf/vscode-spfx-localization/issues/24): Fix for unwanted spaces in localization keys. Thanks to [Nikolay Belykh](https://github.com/nbelyh).

## [1.7.0] - 2020-12-17

- [#11](https://github.com/estruyf/vscode-spfx-localization/issues/11): Support for comments and timestamps added to the CSV export. Thanks to [Nikolay Belykh](https://github.com/nbelyh).

## [1.6.0] - 2020-11-23

- [#12](https://github.com/estruyf/vscode-spfx-localization/issues/12): Support for adding keys in alphabetical order. Thanks to [Nikolay Belykh](https://github.com/nbelyh).

## [1.5.1] - 2020-11-20

- [#5](https://github.com/estruyf/vscode-spfx-localization/issues/5): Fix to support WSL thanks to [Nikolay Belykh](https://github.com/nbelyh)

## [1.5.0] - 2020-11-19

- Added support for `UTF8 BOM` marker option thanks to [Nikolay Belykh](https://github.com/nbelyh)

## [1.4.0] - 2020-09-07

**Enhancements** 

- [#6](https://github.com/estruyf/vscode-spfx-localization/issues/6): Updating MyStrings.d.ts when importing CSV

## [1.3.0] - 2019-08-07

**Enhancements**

- Added support for TypeScript locale files and the ability to specify the locale file extension preference

**Fixes**

- Translations labels not included into the CSV [#1](https://github.com/estruyf/vscode-spfx-localization/issues/1)

## [1.2.2] - 2019-02-06

**Fixes**

- Fixed an issue where the localization label with double quotes usage was incorrectly created

## [1.2.1] - 2018-09-25

**Fixes**

- Performance improvements by temporarily disabling the hover provider

## [1.2.0] - 2018-08-22

**Enhancements**

- Added the ability to export localization labels to a CSV file
- Added the ability to import localization labels from a CSV file
- Context actions should only be shown in SPFx projects

## [1.1.0] - 2018-07-31

**Enhancements**

- Added the localization actions to the editor context menu

## [1.0.0] - 2018-07-13

**Enhancements**

- Added hover provider to show the resource values of the hovered string

## [0.0.5] - 2018-06-14

**Enhancements**

- Removed information messages when localization keys were added

## [0.0.4] - 2018-06-08

**Enhancements**

- Added a new command to let you create a localized key without selecting text
- Code optimizations

## [0.0.3] - 2018-06-07

**Enhancements**

- Remove single and double quotes if string starts and ends with it
- Added the option which allows you to place the new key inside curly brackets `{}`
- Checks if key was already defined

**Fixes**

- Fixes spaces on beginning of the next line where key was added

## [0.0.2] - 2018-06-07

**Fixes**

- Changed `fs.readDirSync` to use the `vscode` api
- Fixed an issue where only the `lib` files were modified instead of the `src` files
- Added a check to see if the `strings` import is already present in the current file

## [0.0.1] - 2018-06-06
- Initial BETA release