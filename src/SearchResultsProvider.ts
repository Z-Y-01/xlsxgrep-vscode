import * as vscode from "vscode";
import * as xlsx from "xlsx";
import * as path from "path";
import { ISearchData } from "./Interfaces";
import {
    SearchResultBookItem,
    SearchResultCellItem,
    SearchResultSheetItem,
} from "./SearchResultItem";

type SearchResultTreeItem =
    | SearchResultBookItem
    | SearchResultCellItem
    | SearchResultSheetItem;

export class SearchResultsProvider
    implements vscode.TreeDataProvider<SearchResultTreeItem>
{
    public static readonly viewType = "xlsxgrep.resultsView";
    private _onDidChangeTreeData: vscode.EventEmitter<
        SearchResultTreeItem | undefined | null | void
    > = new vscode.EventEmitter<
        SearchResultTreeItem | undefined | null | void
    >();
    readonly onDidChangeTreeData: vscode.Event<
        SearchResultTreeItem | undefined | null | void
    > = this._onDidChangeTreeData.event;

    private matchCount: number = 0;
    private books: SearchResultBookItem[] = [];

    public getTreeItem(element: SearchResultTreeItem): vscode.TreeItem {
        return element;
    }

    public getChildren(
        element?: SearchResultTreeItem
    ): vscode.ProviderResult<SearchResultTreeItem[]> {
        if (!element) {
            return this.books;
        }

        if (element instanceof SearchResultBookItem) {
            return element.sheetItems;
        }

        if (element instanceof SearchResultSheetItem) {
            return element.cellItems;
        }

        return [];
    }

    public async runSearch(data: ISearchData) {
        const { targetVal } = data;
        const startOutput = `Start Searching for ${targetVal} ...`;
        vscode.window.showInformationMessage(startOutput);
        console.log(startOutput);

        this.matchCount = 0; // 清空上一次搜索结果
        this._onDidChangeTreeData.fire();

        // 1. 确定搜索范围
        let uris: vscode.Uri[] | undefined = await this._getSearchUris(data);
        if (uris === undefined || uris.length <= 0) {
            vscode.window.showWarningMessage(
                "Please check if there is a .xlsx file in the current workspace."
            );
            return;
        }

        // 2. 执行核心搜索
        vscode.window.withProgress(
            {
                location: { viewId: SearchResultsProvider.viewType },
                title: "XLSXGrep: Searching...",
                cancellable: true,
            },
            async (process, token) => {
                for (const uri of uris) {
                    if (token.isCancellationRequested) {
                        break;
                    }

                    this._processSearchInExcelFile(uri.fsPath, data);
                }
            }
        );

        // 3. 通知 UI 刷新
        this._onDidChangeTreeData.fire();

        let endOutput = `Search complete, found ${this.matchCount} matches.`;
        vscode.window.showInformationMessage(endOutput);
        console.log(endOutput);
    }

    private async _getSearchUris(data: ISearchData) {
        const filePattern: string | undefined = data.filePattern;
        const bCheckFilePattern: boolean =
            filePattern !== undefined && filePattern.length > 0;
        if (!data.bOnlyActiveFiles) {
            return await vscode.workspace.findFiles(
                bCheckFilePattern ? `**/*${filePattern}*.xlsx` : "**/*.xlsx"
            );
        }

        const tabs = vscode.window.tabGroups.activeTabGroup.tabs;
        if (tabs === undefined || tabs.length <= 0) {
            return;
        }

        let uris: vscode.Uri[] = [];
        for (const targetTab of tabs) {
            const label = targetTab.label;
            if (
                label.match(/\.xlsx$/i) &&
                (!bCheckFilePattern ||
                    (filePattern && label.includes(filePattern)))
            ) {
                let uri = (targetTab.input as any)?.uri;
                uri && uris.push(uri);
            }
        }

        return uris;
    }

    private _processSearchInExcelFile(
        filePath: string,
        data: ISearchData
    ): void {
        try {
            const workbook = xlsx.readFile(filePath);
            const fileName = path.basename(filePath);
            const sheetItems: SearchResultSheetItem[] = [];
            for (const sheetName of workbook.SheetNames) {
                const cellItems: SearchResultCellItem[] = [];
                // 将表格转为二维数组，{header: 1} 确保能拿到所有单元格
                const rows: any[][] = xlsx.utils.sheet_to_json(
                    workbook.Sheets[sheetName],
                    {
                        header: 1,
                        defval: "",
                    }
                );
                for (let r = 0; r < rows.length; r++) {
                    for (let c = 0; c < rows[r].length; c++) {
                        const cellValue = String(rows[r][c]);
                        if (this._checkMatch(cellValue, data)) {
                            cellItems.push(
                                new SearchResultCellItem({
                                    path: filePath,
                                    fileName: fileName,
                                    sheet: sheetName,
                                    row: r + 1, // Excel 行号从 1 开始
                                    col: this._indexToColName(c), // 转换索引为 A, B, C...
                                    cellContent: cellValue,
                                    rowContent: String(rows[r]),
                                })
                            );
                            this.matchCount += 1;
                        }
                    }
                }

                if (cellItems.length > 0) {
                    const sheetItem = new SearchResultSheetItem(
                        sheetName,
                        cellItems.length,
                        filePath,
                        cellItems
                    );
                    sheetItems.push(sheetItem);
                }
            }

            if (sheetItems.length > 0) {
                this.books.push(
                    new SearchResultBookItem(fileName, filePath, sheetItems)
                );
            }
        } catch (e) {
            const errorOutput = `Unable to read file: ${filePath}. `;
            vscode.window.showInformationMessage(errorOutput);
            console.error(errorOutput, e);
        }
    }

    private _checkMatch(cellValue: string, data: ISearchData): boolean {
        let { targetVal, isCaseMatch, isWholeMatch, TargetRegexExp } = data;
        if (TargetRegexExp) {
            return TargetRegexExp.test(cellValue);
        }

        // 处理大小写
        if (!isCaseMatch) {
            targetVal = targetVal.toLowerCase();
            cellValue = cellValue.toLowerCase();
        }

        // 处理全字匹配 || 包含匹配
        return isWholeMatch
            ? cellValue === targetVal
            : cellValue.includes(targetVal);
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
