import * as vscode from 'vscode';
import * as fs from 'fs';
import * as path from 'path';
import { ISearchData } from './Interfaces';

export class SearchWebviewProvider implements vscode.WebviewViewProvider {
    public static readonly viewType = 'xlsxgrep.searchView';

    constructor(private readonly _extensionUri: vscode.Uri, private readonly _onSearch: (data: ISearchData) => void) {}

    public resolveWebviewView(webviewView: vscode.WebviewView) {
        webviewView.webview.options = { enableScripts: true };

        // 设置 HTML 内容
        webviewView.webview.html = this._getHtmlForWebview(webviewView.webview);

        // 接收来自 Webview 的消息
        webviewView.webview.onDidReceiveMessage((data: any) => this._processSearchInput(data));
    }

    private _getHtmlForWebview(_webview: vscode.Webview): string {
        return fs.readFileSync(path.join(this._extensionUri.fsPath, '\\src\\html\\SearchWebview.html'), 'utf8');
    }

    private _processSearchInput(data: any){
        if (data.command !== 'search') {
                return;
            }

            const targetVal: string = data.targetVal ?? "";
            if(targetVal.length <= 0)
            {
                vscode.window.showInformationMessage(`Please check if your input is empty.`);
                return;
            }

            let searchData:ISearchData = {
                targetVal: targetVal,
                filePattern: data.filePattern,
                isWholeMatch: data.isWholeExactMatch ?? false,
                isCaseMatch: data.isCaseExactMatch ?? false,
                bOnlyActiveFiles: data.bOnlyActiveFiles ?? false,
            };
            if (data.bUseRegex)
            {
                try{
                    // 'i' 表示忽略大小写
                    const flags = searchData.isCaseMatch ? '' : 'i';
                    let pattern = targetVal;

                    // 如果是全字匹配，包裹边界符
                    if (searchData.isWholeMatch) {
                        pattern = `^${targetVal}$`;
                    }

                    searchData.TargetRegexExp = new RegExp(pattern, flags);
                }
                catch(e){
                    vscode.window.showInformationMessage(`Please check if the regular expression for your search content is entered correctly.`);
                    return;
                }
            }

            this._onSearch(searchData);
    }
}