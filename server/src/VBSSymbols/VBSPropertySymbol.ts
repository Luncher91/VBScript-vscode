import * as ls from 'vscode-languageserver';
import { VBSSymbol } from "./VBSSymbol";

export class VBSPropertySymbol extends VBSSymbol {
	public GetLsName(): string {
		if(this.args != "")
			return this.type + " " + this.name + " (" + this.args + ")";
		else
			return this.type + " " + this.name;
	}

	public GetLsKind(): ls.SymbolKind {
		return ls.SymbolKind.Property;
	}
}