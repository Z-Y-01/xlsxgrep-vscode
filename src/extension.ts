import * as vscode from 'vscode';
import { SearchWebviewProvider } from './SearchWebviewProvider';
import { ISearchData } from './Interfaces';

export function activate(context: vscode.ExtensionContext) {
	console.log('xlsxgrep extension is now active');

	// Register a fallback TreeDataProvider for scenarios where the view was moved to Explorer
	const openSearchCmd = vscode.commands.registerCommand('xlsxgrep.openSearchView', async () => {
		vscode.window.showInformationMessage('Opened XLSXgrep activity bar. Please click the "XLSX Search" view to open the panel.');
	});
	context.subscriptions.push(openSearchCmd);

    // 注册 Webview 视图
	const WebviewProvider = new SearchWebviewProvider(context.extensionUri, (data: ISearchData) => _onSearch(data));
    context.subscriptions.push(
        vscode.window.registerWebviewViewProvider(SearchWebviewProvider.viewType, WebviewProvider)
    );
}

export function deactivate() {
	console.log('xlsxgrep extension is deactive');
}

export function _onSearch(data: ISearchData){
	const {targetVal, isCaseMatch, isWholeMatch} = data;
	const startOutput = `Searching for: ${targetVal}, isWholeExactMatch: ${isWholeMatch}, isCaseExactMatch: ${isCaseMatch}`;
    vscode.window.showInformationMessage(startOutput);
	console.log(startOutput);
}