import * as fs from 'fs';


export function activate_winax() {
    var winax = require('winax'); // required to define ActiveXObject
}

enum SourceType { Excel, M };

interface IDisposable {
    dispose(): void;
}

export class ExcelRegistry implements IDisposable {
    excel: Excel.Application | undefined;
    logger: (arg0: string) => void;

    constructor(logger = function (msg: string) { }) {
        this.logger = logger;
    }

    isInitialized(): boolean {
        return this.excel !== undefined;
    }

    getExcel(): Excel.Application {
        if (!this.excel || !this.excel.Workbooks) {
            this.logger("Create new Excel application");
            this.excel = new ActiveXObject('Excel.Application');
            (this.excel as any).ShowStartupDialog = false;
            this.excel.Visible = true;
        }
        return this.excel;
    }

    dispose(): void {
        if (this.isInitialized() && this.getExcel().Workbooks.Count === 0) {
            this.getExcel().Quit();
        }
    }

    __getWorkbookFromCache(filename: string): Excel.Workbook | undefined {
        if (this.isInitialized()) {
            let excel = this.getExcel();
            for (let i = 1; i <= excel.Workbooks.Count; i++) {
                let wb = excel.Workbooks.Item(i);
                if (wb.FullName === filename) {
                    this.logger("Workbook is already open. Retrieve it.");
                    return wb;
                }
            }
        }
        return undefined;
    }

    open(filename: string): Excel.Workbook {
        let wb: Excel.Workbook | undefined = this.__getWorkbookFromCache(filename);
        if (!wb) {
            this.logger("Open workbook with Excel.")
            wb = this.getExcel().Workbooks.Open(filename);
        }
        return wb;
    }

    close(filename: string, saveChanges: boolean): void {
        let wb: Excel.Workbook | undefined = this.__getWorkbookFromCache(filename);
        if (wb) {
            this.logger("Close workbook " + filename);
            wb.Close(saveChanges);
            wb = undefined;
        }

        if (this.excel && this.excel.Workbooks.Count !== undefined && this.excel.Workbooks.Count === 0) {
            this.logger("All workbooks are closed. Close Excel also.");
            this.getExcel().DisplayAlerts = false;
            this.getExcel().Quit();
            this.excel = undefined;
        }
    }
}



export class PowerQueryMCodeReader implements IDisposable {
    queries: Map<string, string>;
    excelFileName!: string;
    pqmFileName: string;
    pqmFolderName: string="";
    sourceType: SourceType;
    excelRegistry: ExcelRegistry;
    readonly delimiter1: string = "//######";
    readonly delimiter2: string = "## This is delimiter. Dont remove it\n";

    constructor(fileName: string, excelRegistry: ExcelRegistry) {
        if (!fs.existsSync(fileName)) {
            throw new Error("File not found ${fileName}");
        }
        this.excelRegistry = excelRegistry;
        this.queries = new Map();

        let xls_regexp = new RegExp("\\.xls.$");
        let m_regexp = new RegExp("\\.m$");

        if (fileName.toLowerCase().match(xls_regexp)) {
            this.sourceType = SourceType.Excel;
            this.pqmFileName = fileName.toLowerCase().replace(xls_regexp, ".m");
            this.pqmFolderName = fileName.toLowerCase().replace(xls_regexp, "");
            this.excelFileName = fileName;
        } else if (fileName.toLowerCase().match(m_regexp)) {
            this.sourceType = SourceType.M;
            this.pqmFileName = fileName;
            let xlsXname = fileName.toLowerCase().replace(m_regexp, ".xlsx");
            let xlsMname = fileName.toLowerCase().replace(m_regexp, ".xlsm");
            if (fs.existsSync(xlsMname) && fs.existsSync(xlsXname)) {
                throw new Error("xlsX and xlsM files with the same names " +
                    "exist simultaniously. I dont know which to update");
            } else if (fs.existsSync(xlsXname)) {
                this.excelFileName = xlsXname;
            } else if (fs.existsSync(xlsMname)) {
                this.excelFileName = xlsMname;
            }
        } else {
            throw new Error("Unable to handle format of file " + fileName);
        }
    }

    dispose(): void {
        // pass
    }

    importFromExcel(): void {
        console.log("Initalize from Excel file");
        let workbook: any = this.excelRegistry.open(this.excelFileName);
        let queries = workbook["Queries"];
        if (queries === undefined) {
            throw new Error("Queries attribute of Excel Workbook is undefined. It could happen due to raise of Activation window on Excel startup. Check Excel and try again.");
        }
        for (let i = 1; i <= queries.Count; i++) {
            this.queries.set(queries.Item(i).Name, queries.Item(i).Formula);
        }
        console.log(this.queries.size + " queries imported");
    }

    exportToFile(): void {
        console.log("Save to file");
        let buffer: string[] = [];
        let mFolder: string = this.pqmFileName;

        console.log(`mFolder=${mFolder}`)
        if (!fs.existsSync(mFolder)) {
            fs.mkdirSync(mFolder);
        }

        for (let [name, query] of this.queries) {
            let subQueryFlieName: string = `${this.pqmFolderName}/${name}.m`;
           
            let newLine = this.delimiter1 + name + this.delimiter2 + query + "\n\n";
            // buffer.push(newLine);
            fs.writeFileSync(subQueryFlieName, newLine), "utf8";
        }
        // fs.writeFileSync(this.pqmFileName, buffer.join("")), "utf8";
    }

    importFromFile(): void {
        let content: string = fs.readFileSync(this.pqmFileName, "utf8");

        let queries = content.split(this.delimiter1);
        console.log("Read " + queries.length.toString() + " queries");
        this.queries = new Map();
        let delimiter = this.delimiter2.trim();
        for (let queryWithName of queries) {
            let del2pos = queryWithName.indexOf(delimiter);
            if (del2pos == -1) {
                continue;
            }
            let name: string = queryWithName.substring(0, del2pos).trim();
            let query: string = queryWithName.substring(delimiter.length + del2pos, queryWithName.length).trim();
            this.queries.set(name, query.trim());
        }
    }

    exportToExcel(): void {
        console.log("Save to Excel file");
        let workbook: any = this.excelRegistry.open(this.excelFileName);
        let excelQueries = workbook["Queries"];
        if (excelQueries === undefined) {
            throw new Error("Worbook.Queries is undefined. Unable to import. This could be due to Excel pop-up windows.");
        }
        // make a copy of a queries map
        let queriesCopy: Map<string, string> = new Map(this.queries);

        for (let i = 1; i <= excelQueries.Count; i++) {
            let item = excelQueries.Item(i);
            let name = item.Name;
            if (queriesCopy.has(name)) {
                item.Formula = queriesCopy.get(name);
                queriesCopy.delete(name);
            } else {
                // if query exist in Excel, but missed in M file, then drop it from excel
                excelQueries.Item(i).Delete();
                i--;
            }
        }
        // now I need to create new items for ones thar were not present in Excel
        for (let [name, formula] of queriesCopy) {
            excelQueries.Add(name, formula);
        }
    }
}

if (require.main === module) {
    // pass
}