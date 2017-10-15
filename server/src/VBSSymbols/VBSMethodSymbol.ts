import * as ls from 'vscode-languageserver';
import { VBSSymbol } from "./VBSSymbol";

export class VBSMethodSymbol extends VBSSymbol {
	public GetLsName(): string {
		return this.name + " (" + this.args + ")";
	}
	
	public GetLsKind(): ls.SymbolKind {
		return ls.SymbolKind.Method;
	}
}