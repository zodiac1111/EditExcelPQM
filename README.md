# EditExcelPQM - edit M code of your xlsx in VSCode
Want to export and edit Power Query M code of your xlsx file in VSCode and get it back to xlsx? - Here is the plugin. 

Please note
* Plugin adds 4 items in right button pop-up menu named EEPQM***
* It takes ~10 seconds for plugin to startup and show menu items
* On click M->Excel, it forces Excel to update worksheet queries, but **doesn't save** your file
* In order to save your changes you need either to close Excel manually or click 'EEPQM: Close with saving'
* Close actions from popup **doesn't kill** Excel process, but hides it. Dont know how to kill it fully
* It uses COM api to export queries, so it starts an Excel instance
* Unfortunately, I can only start a new Excel instance. I cant use existing one. So, if you have already opened a spreadsheet with M, I will open it again in separate process. If you try to save, it causes exception
* If Excel shows popups on startup, export fails. You need to close Excel popup manually
* Some users reported that Excel starts a process with visbility=false. Not fixed yet.


## Versions
Get Electron version of your VSCode via Help->About
* 1.1.1 supports Electron 7.2.1, NODE_MODULE_VERSION 75
* 1.1.5 supports Electron 9.2.1, NODE_MODULE_VERSION 80
* 1.1.6 supports Electron 11.3.0, NODE_MODULE_VERSION 85
* 1.1.7 supports Electron 19.0.7, NODE_MODULE_VERSION ??

## Features
* Export all M queries from xlsx/xlsm file to *.m file
* Import queries from *.m file to xlsx/xlsm
* Edit M code in VSCode and run queries in Excel immediately 
* Create new queries and upload them to Excel
* Delete queries from VSCode

## Demo
![Image of demo](images/demo.gif)

## Install to Visual Studio Code
From [VSCode extensions market](https://marketplace.visualstudio.com/items?itemName=AMalanov.editexcelpqm) or manually:
1) Download [vsix file](https://github.com/amalanov/EditExcelPQM/blob/master/editexcelpqm-1.1.6.vsix) from git repo
2) Go to download folder
3) Run in console **code --install-extension /path/to/vsix**

## Known issues
* Unable to fully close Excel - window is closed, but it remains in process manager
* If your Excel shows a popup on startup, plugin is unable to access queries before you close the popup
* On some systems the plugin opens Excel in background mode and I'm not able to do it visible.
* (2020-10-24) After VSCode Electron update plugin fails to start and asks for recompilation of native node modules (winax)

## Requirements for 1.1.7
* VSCode ^1.92.0
* Windows
* MS Excel ^2016 - cause the plugin uses AxtiveXObject to open xlsx and extract data
* It takes ~10 seconds for plugin to startup and show menu items

## How to build it yourself
When Electrorn under VSCode updates, you need to recompile this plugin, because it uses native node module winax to access Windows COM api to interface with MS Excel via OLE. You need to do the following
1) Clone this project from github
2) Install Visual Studio C++ compiller for windows ref:[msc](https://visualstudio.microsoft.com/vs/features/cplusplus/)
3) Run VSCode
4) Go to Help -> About 
5) Remember the Electron version
6) Open file package.json of this project
7) Find task `build_winax_for_vscode` in package.json
8) Put the version after `--target=`. For instance, --target=11.3.0
9) Open terminal in VSCode
10) Run `npm install` to download and install all plugin's depenencies 
11) Run NPM scritp `npm run build_vscode_extension`
12) Run NPM scritp `npm run pack_extension`