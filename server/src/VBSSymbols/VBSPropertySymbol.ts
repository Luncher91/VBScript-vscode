import * as ls from 'vscode-languageserver';
import { VBSSymbol } from "./VBSSymbol";

export class VBSPropertySymbol extends VBSSymbol {
	public GetLsName(): string {
		if(this.args != "")
			return this.type + " " + this.name + " (" + this.args + ")";
		else
			return this.type + " " + this.name;
	}

	public GetLsSymbolKind(): ls.SymbolKind {
		return ls.SymbolKind.Property;
	}

	public GetLsCompletionItem(): ls.CompletionItem {
		let item = ls.CompletionItem.create(this.name);
		if(this.args != null)
			item.documentation = this.visibility + " " + this.type.trim() + " " + this.name + "(" + this.args + ")"
		else
			item.documentation = this.visibility + " " + this.type.trim() + " " + this.name
		item.filterText = this.name;
		item.insertText = this.name;
		item.kind = ls.CompletionItemKind.Property;
		return item;
	}
}