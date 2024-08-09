# Change Log

All notable changes to the "EditExcelPQM" extension will be documented in this file.

## [Unreleased]

## [1.1.7] - 2024-08-09
### Changed
- After VSCode Electron update plugin fails to start. I'vr recompiled native node modules (winax). Recompiled for VSCode 1.73.1, Electrol 19.0.17.

## [1.1.6] - 2021-03-06
### Changed
- After VSCode Electron update plugin fails to start. I'vr recompiled native node modules (winax). Recompiled for VSCode 1.54.1, Electrol 11.3.0, NODE_MODULE_VERSION / ABI (application binnary interface) 85
- Modified build_winax_for_vscode NPM script to comply with a new version of Electron

## [1.1.5] - 2020-10-25
### Fixed
- Unable to open xlsm file

## [1.1.4] - 2020-10-25
### Changed
- After VSCode Electron update plugin fails to start and asks for recompilation of native node modules (winax). Recompiled for VSCode 1.50.1
- Change popup items labels after user feedback

## [1.1.1] - 2020-08-08
### Added
- After you delete query from M file and export it to xlsx, a deleted query is dropped in xlsx
- Context menue now has an option to close excel without saving
### Changed
- All items of context menue have EEPQM prefix

## [1.0.0] - 2020-07-26
### Added
- Export all M queries from xlsx/xlsm file to *.m file
- Import queries from *.m file to xlsx/xlsm
- Edit M code in VSCode and run queries in Excel immediately
- Create new queries in VSCode and upload them to Excel
