import * as vscode from 'vscode';
import { SearchWebviewProvider } from './SearchWebviewProvider';
import { SearchResultsProvider } from './SearchResultsProvider';
import { ISearchData, ISearchResult } from './Interfaces';

export function activate(context: vscode.ExtensionContext) {
	console.log('xlsxgrep extension is now active');

    const resultsProvider = new SearchResultsProvider();
	const WebviewProvider = new SearchWebviewProvider(context.extensionUri, (data: ISearchData) => resultsProvider?.runSearch(data));
    context.subscriptions.push(
		vscode.window.registerTreeDataProvider(SearchResultsProvider.viewType, resultsProvider),
        vscode.window.registerWebviewViewProvider(SearchWebviewProvider.viewType, WebviewProvider)
    );
}

export function deactivate() {
	console.log('xlsxgrep extension is deactive');
}
