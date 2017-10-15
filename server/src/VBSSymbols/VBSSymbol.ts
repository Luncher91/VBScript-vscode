import * as ls from 'vscode-languageserver';

export class VBSSymbol {
	public visibility: string = "";
	public name: string = "";
	public type: string = "";
	public args: string = "";
	public symbolRange: ls.Range = null;
	public nameLocation: ls.Location = null;
	
	public parentName: string = "";

	public GetLsName(): string {
		return this.name;
	}

	public GetLsSymbolKind(): ls.SymbolKind {
		// I do not know any better value to return here - I liked to have something like ls.SymbolKind.UNKNOWN
		return ls.SymbolKind.File;
	}

	public GetLsCompletionItem(): ls.CompletionItem {
		let item = ls.CompletionItem.create(this.name);
		item.filterText = this.name;
		item.insertText = this.name;
		item.kind = ls.CompletionItemKind.Text;
		return item;
	}
	
	public static GetLanguageServerSymbols(symbols: VBSSymbol[]): ls.SymbolInformation[] {
		let lsSymbols: ls.SymbolInformation[] = [];

		symbols.forEach(symbol => {
			let lsSymbol: ls.SymbolInformation = ls.SymbolInformation.create(
				symbol.GetLsName(),
				symbol.GetLsSymbolKind(),
				symbol.symbolRange,
				symbol.nameLocation.uri,
				symbol.parentName
			);
			lsSymbols.push(lsSymbol);
		});

		return lsSymbols;
	}

	public static GetLanguageServerCompletionItems(symbols: VBSSymbol[]): ls.CompletionItem[] {
		let completionItems: ls.CompletionItem[] = [];

		symbols.forEach(symbol => {
			let lsItem = symbol.GetLsCompletionItem();
			completionItems.push(lsItem);
		});

		return completionItems;
	}
}