import * as vscode from 'vscode';
import { ISearchResult } from './Interfaces';

export class SearchResultBookItem extends vscode.TreeItem {
    public readonly isFile = true;
    constructor(
        public readonly label: string, 
        public readonly path: string, 
        public readonly results: ISearchResult[]
    ) {
        super(label, vscode.TreeItemCollapsibleState.Expanded);
        this.resourceUri = vscode.Uri.file(path); 
        this.description = `${results.length} matches`;
        this.iconPath = new vscode.ThemeIcon('files');
    }
}

export class SearchResultCellItem extends vscode.TreeItem {
    constructor(public readonly result: ISearchResult) {
        super(result.cellContent, vscode.TreeItemCollapsibleState.None);
        this.description = `[${result.sheet}] $${result.col}$${result.row}, RowContent: ${result.rowContent}`;
        this.tooltip = `Value: ${result.cellContent}`;
        this.iconPath = new vscode.ThemeIcon('symbol-field');
        
        // 点击子节点打开文件
        this.command = {
            command: 'vscode.open',
            title: 'Open File',
            arguments: [vscode.Uri.file(result.path)]
        };
    }
}
