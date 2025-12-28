import * as vscode from 'vscode';
import * as fs from 'fs';
import * as path from 'path';

export class SearchWebviewProvider implements vscode.WebviewViewProvider {
    public static readonly viewType = 'xlsxgrep.searchView';

    constructor(private readonly _extensionUri: vscode.Uri) {}

    resolveWebviewView(webviewView: vscode.WebviewView) {
        webviewView.webview.options = { enableScripts: true };

        // 设置 HTML 内容
        webviewView.webview.html = this._getHtmlForWebview(webviewView.webview);

        // 接收来自 Webview 的消息
        webviewView.webview.onDidReceiveMessage(data => {
            if (data.command !== 'search') {
                return;
            }

            const searchValue = data.text;
            const isWholeMatch = data.isWholeExactMatch;
            const isCaseMatch = data.isCaseExactMatch;
            const  output: string = `Searching for: ${searchValue}, isWholeExactMatch: ${isWholeMatch}, isCaseExactMatch: ${isCaseMatch}`;
            vscode.window.showInformationMessage(output);
        });
    }

    private _getHtmlForWebview(webview: vscode.Webview): string {
        return fs.readFileSync(path.join(this._extensionUri.fsPath, '\\src\\SearchWebview.html'), 'utf8');
    }
}