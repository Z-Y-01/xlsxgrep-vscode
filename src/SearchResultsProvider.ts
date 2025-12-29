import * as vscode from 'vscode';
import * as xlsx from 'xlsx';
import * as path from "path";
import { ISearchResult, ISearchData } from './Interfaces';
import {SearchResultBookItem, SearchResultCellItem, SearchResultSheetItem } from './SearchResultItem';

type SearchResultTreeItem = SearchResultBookItem | SearchResultCellItem | SearchResultSheetItem;

export class SearchResultsProvider implements vscode.TreeDataProvider<SearchResultTreeItem> {
    public static readonly viewType = 'xlsxgrep.resultsView';
    private _onDidChangeTreeData: vscode.EventEmitter<SearchResultTreeItem | undefined | null | void> = new vscode.EventEmitter<SearchResultTreeItem | undefined | null | void>();
    readonly onDidChangeTreeData: vscode.Event<SearchResultTreeItem | undefined | null | void> = this._onDidChangeTreeData.event;

    private filePathToSheetResults: Map<string, Map<string, ISearchResult[]>> = new Map();

    public getTreeItem(element: SearchResultTreeItem): vscode.TreeItem {
        return element;
    }

    public getChildren(element?: SearchResultTreeItem): vscode.ProviderResult<SearchResultTreeItem[]> {
        if (!element) {
            return this._getTreeViewBookItems();
        }

        if (element instanceof SearchResultBookItem) {
            return this._getTreeViewSheetItems(element.filePath);
        }

        if (element instanceof SearchResultSheetItem) {
            return this._getTreeViewCellItem(element.filePath, element.sheetName);
        }

        return [];
    }

    public async runSearch(data: ISearchData){
        const {targetVal} = data;
        const startOutput = `Start Searching for: ${targetVal}...`;
        vscode.window.showInformationMessage(startOutput);
        console.log(startOutput);

        this.filePathToSheetResults.clear(); // 清空上一次搜索结果
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
        }, async (process, token) => {
            for(const uri of uris){
                if (token.isCancellationRequested) {
                    break; 
                }

                const fsPath = uri.fsPath;
                const sheetResults = this._searchInExcelFile(fsPath, data);
                if (sheetResults) {
                    this.filePathToSheetResults.set(fsPath, sheetResults);
                }
            }
        });

        // 3. 通知 UI 刷新
        this._onDidChangeTreeData.fire();

        let endOutput = `Search complete, found ${this.filePathToSheetResults.values()} matches.`;
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

    private _searchInExcelFile(filePath: string, data: ISearchData):  Map<string, ISearchResult[]> | undefined {
        const sheetNameToResults: Map<string, ISearchResult[]> = new Map();
        try {
            const workbook = xlsx.readFile(filePath);
            const fileName = path.basename(filePath);
            for (const sheetName of workbook.SheetNames) {
                const sheet = workbook.Sheets[sheetName];
                const results = sheetNameToResults.get(sheetName) ?? [];
                // 将表格转为二维数组，{header: 1} 确保能拿到所有单元格
                const rows: any[][] = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: "" });
                for (let r = 0; r < rows.length; r++) {
                    for (let c = 0; c < rows[r].length; c++) {
                        const cellValue = String(rows[r][c]);
                        if (this._checkMatch(cellValue, data)) {
                            results.push({
                                path: filePath,
                                fileName: fileName,
                                sheet: sheetName,
                                row: r + 1, // Excel 行号从 1 开始
                                col: this._indexToColName(c), // 转换索引为 A, B, C...
                                cellContent: cellValue,
                                rowContent: String(rows[r]),
                            });
                        }
                    }
                }

                results.length > 0 && sheetNameToResults.set(sheetName, results);
            }
        } catch (e) {
            const errorOutput = `Unable to read file: ${filePath}. `;
            vscode.window.showInformationMessage(errorOutput);
            console.error(errorOutput, e);
        }

        return sheetNameToResults;
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

    private _getTreeViewBookItems() {
        // 返回所有文件节点
        const files: SearchResultBookItem[] = [];
        this.filePathToSheetResults.forEach((sheetResults, filePath) => {
            if (sheetResults.size > 0) {
                files.push(new SearchResultBookItem(path.basename(filePath), filePath));
            }
        });

        return files;
    }

    private _getTreeViewSheetItems(filePath: string) {
        const sheetResults = this.filePathToSheetResults.get(filePath);
        if (!sheetResults)
        {
            return [];
        }

        const sheets: SearchResultSheetItem[] = [];
        sheetResults.forEach((results, sheetName) => {
            if (results.length > 0) 
            {
                sheets.push(new SearchResultSheetItem(sheetName, results.length, filePath));
            }
        });
        return sheets;
    }

    private _getTreeViewCellItem(filePath: string, sheetName: string) {
        const sheetResults = this.filePathToSheetResults.get(filePath);
        if (!sheetResults)
        {
            return [];
        }

        const results = sheetResults.get(sheetName);
        if (!results)
        {
            return [];
        }

        return results.map((v)=> new SearchResultCellItem(v));
    }

}