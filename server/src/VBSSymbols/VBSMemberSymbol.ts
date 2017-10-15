import * as ls from 'vscode-languageserver';
import { VBSSymbol } from "./VBSSymbol";

export class VBSMemberSymbol extends VBSSymbol {
	public GetLsKind(): ls.SymbolKind {
		return ls.SymbolKind.Field;
	}
}