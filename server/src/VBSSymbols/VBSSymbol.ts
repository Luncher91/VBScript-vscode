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

	public GetLsKind(): ls.SymbolKind {
		// I do not know any better value to return here - I liked to have something like ls.SymbolKind.UNKNOWN
		return ls.SymbolKind.File;
	}
	
	public static GetLanguageServerSymbols(symbols: VBSSymbol[]): ls.SymbolInformation[] {
		let lsSymbols: ls.SymbolInformation[] = [];

		symbols.forEach(symbol => {
			let lsSymbol: ls.SymbolInformation = ls.SymbolInformation.create(
				symbol.GetLsName(),
				symbol.GetLsKind(),
				symbol.symbolRange,
				symbol.nameLocation.uri,
				symbol.parentName
			);
			lsSymbols.push(lsSymbol);
		});

		return lsSymbols;
	}
}