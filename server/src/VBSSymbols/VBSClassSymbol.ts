import * as ls from 'vscode-languageserver';
import { VBSSymbol } from "./VBSSymbol";

export class VBSClassSymbol extends VBSSymbol {
	public GetLsSymbolKind(): ls.SymbolKind {
		return ls.SymbolKind.Class;
	}

	public GetLsCompletionItem(): ls.CompletionItem {
		let item = ls.CompletionItem.create(this.name);
		item.filterText = this.name;
		item.insertText = this.name;
		item.kind = ls.CompletionItemKind.Class;
		return item;
	}
}