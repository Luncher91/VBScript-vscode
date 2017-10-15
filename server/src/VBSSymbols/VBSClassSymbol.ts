import * as ls from 'vscode-languageserver';
import { VBSSymbol } from "./VBSSymbol";

export class VBSClassSymbol extends VBSSymbol {
	public GetLsKind(): ls.SymbolKind {
		return ls.SymbolKind.Class;
	}
}