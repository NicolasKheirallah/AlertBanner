# Change Log

## [1.2.4] - 2019-05-30

**Fix**

- Escaping folder and gulp paths on the `win32` platform

## [1.2.3] - 2019-05-16

**Fix**

- Reverting fix in `1.2.2` to find a better way to escape the characters

## [1.2.2] - 2019-05-16

**Fix**

- Path escaping for Windows when `(` or `)` characters are used in the gulp path

## [1.2.1] - 2019-04-26

**Fix**

- Path escaping for Windows when `(` or `)` characters are used in the root folder path

## [1.2.0] - 2019-02-07

**Enhancement**

- Added support for multi-root workspaces

## [1.1.1] - 2018-09-04

**Fix**

- Packaging tasks did not correctly initiate its dependent tasks

## [1.1.0] - 2018-09-04

**Enhancement**

- Added `gulp clean` to the `Create release package` command

**Fixes**

- Fix for `undefined` task type

## [1.0.0] - 2018-08-08

**Enhancement**

- Allow to provide arguments to tasks you want to run

## [0.0.2] - 2018-08-02

**Enhancement**

- Added the ability to pick and run one of the available tasks

## [0.0.1] - 2018-08-01

- Initial release