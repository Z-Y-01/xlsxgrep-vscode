import * as vscode from 'vscode';
import * as fs from 'fs';
import * as path from 'path';
import { ISearchData } from './Interfaces';

export class SearchWebviewProvider implements vscode.WebviewViewProvider {
    public static readonly viewType = 'xlsxgrep.searchView';

    constructor(private readonly _extensionUri: vscode.Uri, private readonly _onSearch: (data: ISearchData) => void) {}

    resolveWebviewView(webviewView: vscode.WebviewView) {
        webviewView.webview.options = { enableScripts: true };

        // 设置 HTML 内容
        webviewView.webview.html = this._getHtmlForWebview(webviewView.webview);

        // 接收来自 Webview 的消息
        webviewView.webview.onDidReceiveMessage(data => {
            if (data.command !== 'search') {
                return;
            }

            this._onSearch(
                {
                    targetVal: data.text ?? "",
                    isWholeMatch: data.isWholeExactMatch ?? false,
                    isCaseMatch: data.isCaseExactMatch ?? false,
                }
            );
        });
    }

    private _getHtmlForWebview(webview: vscode.Webview): string {
        return fs.readFileSync(path.join(this._extensionUri.fsPath, '\\src\\html\\SearchWebview.html'), 'utf8');
    }
}