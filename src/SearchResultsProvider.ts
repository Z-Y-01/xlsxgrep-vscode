import * as vscode from 'vscode';
import * as xlsx from 'xlsx';
import * as path from "path";
import { ISearchResult, ISearchData } from './Interfaces';

export class SearchResultsProvider implements vscode.TreeDataProvider<ISearchResult> {
    public static readonly viewType = 'xlsxgrep.resultsView';
    private _onDidChangeTreeData: vscode.EventEmitter<ISearchResult | undefined | null | void> = new vscode.EventEmitter<ISearchResult | undefined | null | void>();
    readonly onDidChangeTreeData: vscode.Event<ISearchResult | undefined | null | void> = this._onDidChangeTreeData.event;

    private results: ISearchResult[] = [];

    getTreeItem(element: ISearchResult): vscode.TreeItem {
        const treeItem = new vscode.TreeItem(element.cellContent); // 单元格内容作为标题
        treeItem.description = `FILE: ${element.fileName}, SHEET: ${element.sheet}, ROW: ${element.row}, COL: ${element.col}`;
        treeItem.tooltip = `Click to open ${element.fileName}.`;
        treeItem.iconPath = new vscode.ThemeIcon('table');
        
        // 如果想要点击跳转
        treeItem.command = {
            command: 'vscode.open',
            title: 'Open File',
            arguments: [vscode.Uri.file(element.path)]
        };
        return treeItem;
    }

    getChildren(_element?: ISearchResult): Thenable<ISearchResult[]> {
        return Promise.resolve(this.results);
    }

    
    async runSearch(data: ISearchData){
        const {targetVal} = data;
        const startOutput = `Start Searching for: ${targetVal}...`;
        vscode.window.showInformationMessage(startOutput);
        console.log(startOutput);

        this.results = []; // 清空上一次搜索结果
        this._onDidChangeTreeData.fire();
        
        // 1. 确定搜索范围
        let uris: vscode.Uri[] | undefined = await this._getSearchUris(data);    
        if(uris === undefined || uris.length <= 0)
        {
            vscode.window.showWarningMessage('Please check if there is a .xlsx file in the current workspace.');
            return;
        }

        // 2. 执行核心搜索
         vscode.window.withProgress({
            location: { viewId: SearchResultsProvider.viewType },
            title: "XLSXGrep: Searching...",
            cancellable: true,
        }, async (_, token) => {
            for(const uri of uris){
                if (token.isCancellationRequested) {
                    break; 
                }

                this.results.push(...this._searchInExcelFile(uri.fsPath, data));
            }
        });

        // 3. 通知 UI 刷新
        this._onDidChangeTreeData.fire();

        let endOutput = `Search complete, found ${this.results.length} matches.`;
        vscode.window.showInformationMessage(endOutput);
        console.log(endOutput);
    }

    private async _getSearchUris(data: ISearchData){
        const filePattern: string | undefined = data.filePattern;
        const bCheckFilePattern: boolean = filePattern !== undefined && filePattern.length > 0;
        if (!data.bOnlyActiveFiles) {
            return await vscode.workspace.findFiles(bCheckFilePattern ? `**/*${filePattern}*.xlsx` : '**/*.xlsx');
        } 

        const tabs = vscode.window.tabGroups.activeTabGroup.tabs;
        if (tabs === undefined || tabs.length <= 0 ) {
           return;
        }

        let uris: vscode.Uri[] = [];
        for (const targetTab of tabs)
        {
            const label = targetTab.label;
            if(label.match(/\.xlsx$/i) && (!bCheckFilePattern ||(filePattern && label.includes(filePattern))))
            {
                let uri = ((targetTab.input as any)?.uri);
                uri && uris.push(uri);
            }
        }

        return uris;
    }

    private _searchInExcelFile(filePath: string, data: ISearchData): ISearchResult[] {
        const fileResults: ISearchResult[] = [];
        try {
            const workbook = xlsx.readFile(filePath);
            const fileName = path.basename(filePath);
            for (const sheetName of workbook.SheetNames) {
                const sheet = workbook.Sheets[sheetName];
                // 将表格转为二维数组，{header: 1} 确保能拿到所有单元格
                const rows: any[][] = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: "" });
                for (let r = 0; r < rows.length; r++) {
                    for (let c = 0; c < rows[r].length; c++) {
                        const cellValue = String(rows[r][c]);
                        if (this._checkMatch(cellValue, data)) {
                            fileResults.push({
                                path: filePath,
                                fileName: fileName,
                                sheet: sheetName,
                                row: r + 1, // Excel 行号从 1 开始
                                col: this._indexToColName(c), // 转换索引为 A, B, C...
                                cellContent: cellValue
                            });
                        }
                    }
                }
            }
        } catch (e) {
            const errorOutput = `Unable to read file: ${filePath}. `;
            vscode.window.showInformationMessage(errorOutput);
            console.error(errorOutput, e);
        }

        return fileResults;
    }

    private _checkMatch(cellValue: string, data: ISearchData): boolean {
        let {targetVal, isCaseMatch, isWholeMatch, TargetRegexExp} = data;
        if (TargetRegexExp) {
            return TargetRegexExp.test(cellValue);
        }

        // 处理大小写
        if (!isCaseMatch) {
            targetVal = targetVal.toLowerCase();
            cellValue = cellValue.toLowerCase();
        }

        // 处理全字匹配 || 包含匹配
        return isWholeMatch ? cellValue === targetVal : cellValue.includes(targetVal);
    }

    private _indexToColName(index: number): string {
        let name = "";
        while (index >= 0) {
            name = String.fromCharCode((index % 26) + 65) + name;
            index = Math.floor(index / 26) - 1;
        }

        return name;
    }

}