import * as vscode from 'vscode';
import { ISearchResult } from './Interfaces';

export class SearchResultBookItem extends vscode.TreeItem {
    constructor(
        public readonly fileName: string, 
        public readonly filePath: string, 
    ) {
        super(fileName, vscode.TreeItemCollapsibleState.Expanded);
        this.resourceUri = vscode.Uri.file(filePath); 
        this.description = ` ${filePath}`;
        this.iconPath = new vscode.ThemeIcon('files');
    }
}

export class SearchResultSheetItem extends vscode.TreeItem {
    constructor(
        public readonly sheetName: string, 
        public readonly matchCount: number,
        public readonly filePath: string,
    ){
        super(sheetName, vscode.TreeItemCollapsibleState.Expanded);
        this.description = ` ${matchCount} matches`;
        this.iconPath = new vscode.ThemeIcon('file');
    }
}

export class SearchResultCellItem extends vscode.TreeItem {
    constructor(public readonly result: ISearchResult) {
        super(result.cellContent, vscode.TreeItemCollapsibleState.None);
        this.description = ` Row: ${result.row}, Col: ${result.col}, RowContent: ${result.rowContent}`;
        this.tooltip = ` Click to open ${result.fileName}.`;
        this.iconPath = new vscode.ThemeIcon('symbol-field');
        
        // 点击子节点打开文件
        this.command = {
            command: 'vscode.open',
            title: 'Open File',
            arguments: [vscode.Uri.file(result.path)]
        };
    }
}
