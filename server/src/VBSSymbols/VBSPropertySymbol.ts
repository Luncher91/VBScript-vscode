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
		
		item.documentation = this.GetDocu();
		item.filterText = this.name;
		item.insertText = this.name;
		item.kind = ls.CompletionItemKind.Property;
		return item;
	}

	private GetDocu(): string {
		let docu = "";
		
		if(this.visibility != null)
			docu += this.visibility.trim();

		if(this.type != null)
			docu += this.type.trim();

		if(this.name != null)
			docu += this.name.trim();

		if(this.args != null)
			docu += "(" + this.args + ")";

		return docu.trim();
	}
}