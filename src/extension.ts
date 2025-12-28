import * as vscode from 'vscode';
import { SearchWebviewProvider } from './SearchWebviewProvider';

export function activate(context: vscode.ExtensionContext) {
	console.log('xlsxgrep extension is now active');

	// Register a fallback TreeDataProvider for scenarios where the view was moved to Explorer
	const openSearchCmd = vscode.commands.registerCommand('xlsxgrep.openSearchView', async () => {
		vscode.window.showInformationMessage('Opened XLSXgrep activity bar. Please click the "XLSX Search" view to open the panel.');
	});
	context.subscriptions.push(openSearchCmd);

    // 注册 Webview 视图
	const WebviewProvider = new SearchWebviewProvider(context.extensionUri);
    context.subscriptions.push(
        vscode.window.registerWebviewViewProvider(SearchWebviewProvider.viewType, WebviewProvider)
    );

	const disposable = vscode.commands.registerCommand('xlsxgrep-vscode.helloWorld', () => {
		vscode.window.showInformationMessage('Hello World from xlsxgrep-vscode!');
	});
	context.subscriptions.push(disposable);
}

export function deactivate() {
	console.log('xlsxgrep extension is deactive');
}