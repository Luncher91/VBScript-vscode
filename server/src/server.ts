/* --------------------------------------------------------------------------------------------
 * Copyright (c) Andreas Lenzen. All rights reserved.
 * Licensed under the MIT License. See License.txt in the project root for license information.
 * ------------------------------------------------------------------------------------------ */
'use strict';

import * as ls from 'vscode-languageserver';
import { VBSSymbol } from "./VBSSymbols/VBSSymbol";
import { VBSMethodSymbol } from './VBSSymbols/VBSMethodSymbol';
import { VBSPropertySymbol } from './VBSSymbols/VBSPropertySymbol';
import { VBSClassSymbol } from './VBSSymbols/VBSClassSymbol';
import { VBSMemberSymbol } from './VBSSymbols/VBSMemberSymbol';
import { VBSVariableSymbol } from './VBSSymbols/VBSVariableSymbol';
import { VBSConstantSymbol } from './VBSSymbols/VBSConstantSymbol';

// Create a connection for the server. The connection uses Node's IPC as a transport
let connection: ls.IConnection = ls.createConnection(new ls.IPCMessageReader(process), new ls.IPCMessageWriter(process));

// Create a simple text document manager. The text document manager
// supports full document sync only
let documents: ls.TextDocuments = new ls.TextDocuments();
// Make the text document manager listen on the connection
// for open, change and close text document events
documents.listen(connection);

// After the server has started the client sends an initialize request. The server receives
// in the passed params the rootPath of the workspace plus the client capabilities.
let workspaceRoot: string;
connection.onInitialize((params): ls.InitializeResult => {
	workspaceRoot = params.rootPath;
	return {
		capabilities: {
			// Tell the client that the server works in FULL text document sync mode
			textDocumentSync: documents.syncKind,
			documentSymbolProvider: true,
			// Tell the client that the server support code complete
			completionProvider: {
				resolveProvider: true
			}
		}
	}
});

// The content of a text document has changed. This event is emitted
// when the text document first opened or when its content has changed.
documents.onDidChangeContent((change: ls.TextDocumentChangeEvent) => {
});

connection.onDidChangeWatchedFiles((changeParams: ls.DidChangeWatchedFilesParams) => {
	for (let i = 0; i < changeParams.changes.length; i++) {
		let event = changeParams.changes[i];

		switch(event.type) {
		 case ls.FileChangeType.Changed:
		 case ls.FileChangeType.Created:
			RefreshDocumentsSymbols(event.uri);
			break;
		case ls.FileChangeType.Deleted:
			symbolCache[event.uri] = null;
			break;
		}
	}
});

// This handler provides the initial list of the completion items.
connection.onCompletion((_textDocumentPosition: ls.TextDocumentPositionParams): ls.CompletionItem[] => {
	return SelectCompletionItems(_textDocumentPosition);
});

connection.onCompletionResolve((complItem: ls.CompletionItem): ls.CompletionItem => {
	return complItem;
});

function GetSymbolsOfDocument(uri: string) : ls.SymbolInformation[] {
	RefreshDocumentsSymbols(uri);
	return VBSSymbol.GetLanguageServerSymbols(symbolCache[uri]);
}

function SelectCompletionItems(textDocumentPosition: ls.TextDocumentPositionParams): ls.CompletionItem[] {
	let symbols = symbolCache[textDocumentPosition.textDocument.uri];

	if(symbols == null) {
		RefreshDocumentsSymbols(textDocumentPosition.textDocument.uri);
		symbols = symbolCache[textDocumentPosition.textDocument.uri];
	}

	let scopeSymbols = GetSymbolsOfScope(symbols, textDocumentPosition.position);
	return VBSSymbol.GetLanguageServerCompletionItems(scopeSymbols);
}

function GetVBSSymbolTree(symbols: VBSSymbol[]) {
	// sort by start positition
	let sortedSymbols: VBSSymbol[] = symbols.sort(function(a: VBSSymbol, b: VBSSymbol){
		let diff = a.symbolRange.start.line - b.symbolRange.start.line;
		
		if(diff != 0)
			return diff;

		return a.symbolRange.start.character - b.symbolRange.start.character;
	});

	let root = new VBSSymbolTree();
	
	for (var i = 0; i < sortedSymbols.length; i++) {
		var symbol = sortedSymbols[i];
		root.InsertIntoTree(symbol);
	}

	return root;
}

function GetSymbolsOfScope(symbols: VBSSymbol[], position: ls.Position): VBSSymbol[] {
	let symbolTree = GetVBSSymbolTree(symbols);
	// bacause of hoisting we will have just a few possible scopes:
	// - file wide
	// - method of file wide
	// - class scope
	// - method or property of class scope
	// get all symbols which are accessable from here (ignore visibility in the first step)

	return symbolTree.FindDirectParent(position).GetAllParentsAndTheirDirectChildren();
}

class VBSSymbolTree {
	parent: VBSSymbolTree = null;
	children: VBSSymbolTree[] = [];
	data: VBSSymbol = null;

	public InsertIntoTree(symbol: VBSSymbol): boolean {
		if(this.data != null && !PositionInRange(this.data.symbolRange, symbol.symbolRange.start))
			return false;

		for (var i = 0; i < this.children.length; i++) {
			var symbolTree = this.children[i];
			if(symbolTree.InsertIntoTree(symbol))
				return true;
		}

		let newTreeNode = new VBSSymbolTree();
		newTreeNode.data = symbol;
		newTreeNode.parent = this;

		this.children.push(newTreeNode);

		return true;
	}

	public FindDirectParent(position: ls.Position): VBSSymbolTree {
		if(this.data != null && !PositionInRange(this.data.symbolRange, position))
			return null;
		
		for (var i = 0; i < this.children.length; i++) {
			let symbolTree = this.children[i];
			let found = symbolTree.FindDirectParent(position);
			if(found != null)
				return found;
		}

		return this;
	}

	public GetAllParentsAndTheirDirectChildren(): VBSSymbol[] {
		let symbols: VBSSymbol[];

		if(this.parent != null)
			symbols = this.parent.GetAllParentsAndTheirDirectChildren();
		else
			symbols = [];
		
		let childSymbols = this.children.map(function(symbolTree) {
			return symbolTree.data;
		});

		return symbols.concat(childSymbols);
	}
}

function PositionInRange(range: ls.Range, position: ls.Position): boolean {
	if(range.start.line > position.line)
		return false;

	if(range.end.line < position.line)
		return false;

	if(range.start.line == position.line && range.start.character >= position.character)
		return false;
		
	if(range.end.line == position.line && range.end.character <= position.character)
		return false;

	return true;
}

let symbolCache: { [id: string] : VBSSymbol[]; } = {};
function RefreshDocumentsSymbols(uri: string) {
	let startTime: number = Date.now();
	let symbolsList: VBSSymbol[] = CollectSymbols(documents.get(uri));
	symbolCache[uri] = symbolsList;
	console.info("Found " + symbolsList.length + " symbols in '" + uri + "': " + (Date.now() - startTime) + " ms");
}

connection.onDocumentSymbol((docParams: ls.DocumentSymbolParams): ls.SymbolInformation[] => {
	return GetSymbolsOfDocument(docParams.textDocument.uri);
});

function CollectSymbols(document: ls.TextDocument): VBSSymbol[] {
	let symbols: Set<VBSSymbol> = new Set<VBSSymbol>();
	let lines = document.getText().split(/\r?\n/g);

	let startMultiLine: number = -1;
	let multiLines: string[] = [];

	for (var i = 0; i < lines.length; i++) {
		let line = lines[i];
		
		let containsComment = line.indexOf("'");
		if(containsComment > -1)
			line = line.substring(0, containsComment);

		if(startMultiLine == -1)
			startMultiLine = i;

		if(line.trim().endsWith("_")) {
			multiLines.push(line.slice(0, -1));
			continue;
		} else {
			multiLines.push(line);
		}

		multiLines = ReplaceStringLiterals(multiLines);
		let statements = SplitStatements(multiLines, startMultiLine);

		statements.forEach(statement => {
			FindSymbol(statement, document.uri, symbols);
		});

		startMultiLine =-1;
		multiLines = [];
	}

	return Array.from(symbols);
}

class MultiLineStatement {
	startCharacter: number = 0;
	startLine: number = -1;
	lines: string[] = [];

	public GetFullStatement(): string {
		return " ".repeat(this.startCharacter) + this.lines.join("");
	}

	public GetPostitionByCharacter(charIndex: number) : ls.Position {
		let internalIndex = charIndex - this.startCharacter;

		for (let i = 0; i < this.lines.length; i++) {
			let line = this.lines[i];
			
			if(internalIndex <= line.length) {
				if(i == 0)
					return ls.Position.create(this.startLine + i, internalIndex + this.startCharacter);
				else
					return ls.Position.create(this.startLine + i, internalIndex);

			}

			internalIndex = internalIndex - line.length;

			if(internalIndex < 0)
				break;
		}

		console.warn("WARNING: cannot resolve " + charIndex + " in me: " + JSON.stringify(this));
		return null;
	}
}

function SplitStatements(lines: string[], startLineIndex: number): MultiLineStatement[] {
	let statement: MultiLineStatement = new MultiLineStatement();
	let statements: MultiLineStatement[] = [];
	let charOffset: number = 0;

	for (var i = 0; i < lines.length; i++) {
		var line = lines[i];
		charOffset = 0;
		let sts: string[] = line.split(":");
		
		if(sts.length == 1) {
			statement.lines.push(sts[0]);
			if(statement.startLine == -1) {
				statement.startLine = startLineIndex + i;
			}
		} else {
			for (var j = 0; j < sts.length; j++) {
				var st = sts[j];
				
				if(statement.startLine == -1)
					statement.startLine = startLineIndex + i;
				
				statement.lines.push(st);
				
				if(j == sts.length-1) {
					break;
				}
				
				statement.startCharacter = charOffset;
				statements.push(statement);
				statement = new MultiLineStatement();

				charOffset += st.length 
					+ 1; // ":"
			}
		}
	}

	if(statement.startLine != -1) {
		statement.startCharacter = charOffset;
		statements.push(statement);
	}

	return statements;
}

function ReplaceStringLiterals(lines:string[]) : string[] {
	let newLines: string[] = [];

	for (var i = 0; i < lines.length; i++) {
		var line = lines[i];
		let stringLiterals = /\"(([^\"]|\"\")*)\"/gi;
		newLines.push(line.replace(stringLiterals, ReplaceBySpaces));
	}

	return newLines;
}

function ReplaceBySpaces(match: string) : string {
	return " ".repeat(match.length);
}

function AddArrayToSet(s: Set<any>, a: any[]) {
	a.forEach(element => {
		s.add(element);
	});
}

function FindSymbol(statement: MultiLineStatement, uri: string, symbols: Set<VBSSymbol>) : void {
	let newSym: VBSSymbol;
	let newSyms: VBSVariableSymbol[] = null;

	if(GetMethodStart(statement, uri)) {
		return;
	}

	newSyms = GetMethodSymbol(statement, uri);
	if(newSyms != null && newSyms.length != 0) {
		AddArrayToSet(symbols, newSyms);
		return;
	}

	if(GetPropertyStart(statement, uri))
		return;

	newSyms = GetPropertySymbol(statement, uri);;
	if(newSyms != null && newSyms.length != 0) {
		AddArrayToSet(symbols, newSyms);
		return;
	}

	if(GetClassStart(statement, uri))
		return;

	newSym = GetClassSymbol(statement, uri);
	if(newSym != null) {
		symbols.add(newSym);
		return;
	}

	newSym = GetMemberSymbol(statement, uri);
	if(newSym != null) {
		symbols.add(newSym);
		return;
	}

	newSyms = GetVariableSymbol(statement, uri);
	if(newSyms != null && newSyms.length != 0) {
		AddArrayToSet(symbols, newSyms);
		return;
	}

	newSym = GetConstantSymbol(statement, uri);
	if(newSym != null) {
		symbols.add(newSym);
		return;
	}
}

let openClassName : string = null;
let openClassStart : ls.Position = ls.Position.create(-1, -1);

class OpenMethod {
	visibility: string;
	type: string;
	name: string;
	argsIndex: number;
	args: string;
	startPosition: ls.Position;
	nameLocation: ls.Location;
	statement: MultiLineStatement;
}

let openMethod: OpenMethod = null;

function GetMethodStart(statement: MultiLineStatement, uri: string): boolean {
	let line = statement.GetFullStatement();

	let rex:RegExp = /^[ \t]*(public[ \t]+|private[ \t]+)?(function|sub)([ \t]+)([a-zA-Z0-9\-\_]+)([ \t]*)(\(([a-zA-Z0-9\_\-, \t(\(\))]*)\))?[ \t]*$/gi;
	let regexResult = rex.exec(line);

	if(regexResult == null || regexResult.length < 6)
		return;

	if(openMethod == null) {
		let leadingSpaces = GetNumberOfFrontSpaces(line);
		let preLength = leadingSpaces + regexResult.index;

		for (var i = 1; i < 6; i++) {
			var resElement = regexResult[i];
			if(resElement != null)
				preLength += resElement.length;
		}

		openMethod = {
			visibility: regexResult[1],
			type: regexResult[2],
			name: regexResult[4],
			argsIndex: preLength + 1, // opening bracket
			args: regexResult[7],
			startPosition: statement.GetPostitionByCharacter(leadingSpaces),
			nameLocation: ls.Location.create(uri, ls.Range.create(
				statement.GetPostitionByCharacter(line.indexOf(regexResult[3])),
				statement.GetPostitionByCharacter(line.indexOf(regexResult[3]) + regexResult[3].length))
			),
			statement: statement
		};
		
		if(openMethod.args == null)
			openMethod.args = "";

		return true;
	} else {
		// ERROR!!! I expected "end function|sub"!
		console.error("ERROR - line " + statement.startLine + " at " + statement.startCharacter + ": 'end " + openMethod.type + "' expected!");
	}

	return false;
}

function GetMethodSymbol(statement: MultiLineStatement, uri: string) : VBSSymbol[] {
	let line: string = statement.GetFullStatement();

	let classEndRegex:RegExp = /^[ \t]*end[ \t]+(function|sub)[ \t]*$/gi;

	let regexResult = classEndRegex.exec(line);

	if(regexResult == null || regexResult.length < 2)
		return null;

	let type = regexResult[1];

	if(openMethod == null) {
		// ERROR!!! I cannot close any method!
		console.error("ERROR - line " + statement.startLine + " at " + statement.startCharacter + ": There is no " + type + " to end!");
		return null;
	}

	if(type.toLowerCase() != openMethod.type.toLowerCase()) {
		// ERROR!!! I expected end function|sub and not sub|function!
		// show the user the error and then go on like it was the right type!
		console.error("ERROR - line " + statement.startLine + " at " + statement.startCharacter + ": 'end " + openMethod.type + "' expected!");
	}

	let range: ls.Range = ls.Range.create(openMethod.startPosition, statement.GetPostitionByCharacter(GetNumberOfFrontSpaces(line) + regexResult[0].trim().length))
	
	let symbol: VBSMethodSymbol = new VBSMethodSymbol();
	symbol.visibility = openMethod.visibility;
	symbol.type = openMethod.type;
	symbol.name = openMethod.name;
	symbol.args = openMethod.args;
	symbol.nameLocation = openMethod.nameLocation;
	symbol.parentName = openClassName;
	symbol.symbolRange = range;

	let parametersSymbol = GetParameterSymbols(openMethod.args, openMethod.argsIndex, openMethod.statement, uri);

	openMethod = null;

	//return [symbol];
	return parametersSymbol.concat(symbol);
}

function ReplaceAll(target: string, search: string, replacement: string): string {
    return target.replace(new RegExp(search, 'g'), replacement);
};

function GetParameterSymbols(args: string, argsIndex: number, statement: MultiLineStatement, uri: string): VBSVariableSymbol[] {
	let symbols: VBSVariableSymbol[] = [];

	if(args == null || args == "")
		return symbols;

	let argsSplitted: string[] = args.split(',');

	for (let i = 0; i < argsSplitted.length; i++) {
		let arg = argsSplitted[i];
		
		let splittedByValByRefName = ReplaceAll(ReplaceAll(arg, "\t", " "), "  ", " ").trim().split(" ");

		let varSymbol:VBSVariableSymbol = new VBSVariableSymbol();
		varSymbol.args = "";
		varSymbol.type = "";
		varSymbol.visibility = "";

		if(splittedByValByRefName.length == 1)
			varSymbol.name = splittedByValByRefName[0].trim();
		else if(splittedByValByRefName.length > 1)
		{
			// ByVal or ByRef
			varSymbol.type = splittedByValByRefName[0].trim();
			varSymbol.name = splittedByValByRefName[1].trim();
		}

		let range = ls.Range.create(
			statement.GetPostitionByCharacter(argsIndex + arg.indexOf(varSymbol.name)),
			statement.GetPostitionByCharacter(argsIndex + arg.indexOf(varSymbol.name) + varSymbol.name.length)
		);
		varSymbol.nameLocation = ls.Location.create(uri, range);
		varSymbol.symbolRange = range;

		symbols.push(varSymbol);
		argsIndex += arg.length + 1; // comma
	}

	return symbols;
}

function GetNumberOfFrontSpaces(line: string): number {
	let counter: number = 0;

	for (var i = 0; i < line.length; i++) {
		var char = line[i];
		if(char == " " || char == "\t")
			counter++;
		else
			break;
	}

	return counter;
}

class OpenProperty {
	visibility: string;
	type: string;
	name: string;
	argsIndex: number;
	args: string;
	startPosition: ls.Position;
	nameLocation: ls.Location;
	statement: MultiLineStatement;
}

let openProperty: OpenProperty = null;

function GetPropertyStart(statement: MultiLineStatement, uri: string) : boolean {
	let line: string = statement.GetFullStatement();

	let propertyStartRegex:RegExp = /^[ \t]*(public[ \t]+|private[ \t]+)?(default[ \t]+)?(property[ \t]+)(let[ \t]+|set[ \t]+|get[ \t]+)([a-zA-Z0-9\-\_]+)([ \t]*)(\(([a-zA-Z0-9\_\-, \t(\(\))]*)\))?[ \t]*$/gi;
	let regexResult = propertyStartRegex.exec(line);

	if(regexResult == null || regexResult.length < 6)
		return null;

	let leadingSpaces = GetNumberOfFrontSpaces(line);
	let preLength = leadingSpaces + regexResult.index;
	
	for (var i = 1; i < 7; i++) {
		var resElement = regexResult[i];
		if(resElement != null)
			preLength += resElement.length;
	}

	if(openProperty == null) {
		openProperty = {
			visibility: regexResult[1],
			type: regexResult[4],
			name: regexResult[5],
			argsIndex: preLength + 1,
			args: regexResult[8],
			startPosition: statement.GetPostitionByCharacter(leadingSpaces),
			nameLocation: ls.Location.create(uri, ls.Range.create(
				statement.GetPostitionByCharacter(line.indexOf(regexResult[5])),
				statement.GetPostitionByCharacter(line.indexOf(regexResult[5]) + regexResult[5].length))
			),
			statement: statement
		};

		if(openProperty.args == null)
			openProperty.args = "";

		return true;
	} else {
		// ERROR!!! I expected "end function|sub"!
		console.error("ERROR - line " + statement.startLine + " at " + statement.startCharacter + ": 'end property' expected!");
	}

	return false;
}

function GetPropertySymbol(statement: MultiLineStatement, uri: string) : VBSSymbol[] {
	let line: string = statement.GetFullStatement();

	let classEndRegex:RegExp = /^[ \t]*end[ \t]+property[ \t]*$/gi;

	let regexResult = classEndRegex.exec(line);

	if(regexResult == null || regexResult.length < 1)
		return null;

	if(openProperty == null) {
		// ERROR!!! I cannot close any property!
		console.error("ERROR - line " + statement.startLine + " at " + statement.startCharacter + ": There is no property to end!");
		return null;
	}

	// range of the whole definition
	let range: ls.Range = ls.Range.create(
		openProperty.startPosition, 
		statement.GetPostitionByCharacter(GetNumberOfFrontSpaces(line) + regexResult[0].trim().length)
	);
	
	let symbol = new VBSPropertySymbol();
	symbol.visibility = "";
	symbol.type = openProperty.type;
	symbol.name = openProperty.name;
	symbol.args = openProperty.args;
	symbol.symbolRange = range;
	symbol.nameLocation = openProperty.nameLocation;
	symbol.parentName = openClassName;
	symbol.symbolRange = range;

	let parametersSymbol = GetParameterSymbols(openProperty.args, openProperty.argsIndex, openProperty.statement, uri);

	openProperty = null;

	return parametersSymbol.concat(symbol);
}

function GetMemberSymbol(statement: MultiLineStatement, uri: string) : VBSMemberSymbol {
	let line: string = statement.GetFullStatement();

	let memberStartRegex:RegExp = /^[ \t]*(public[ \t]+|private[ \t]+)([a-zA-Z0-9\-\_]+)[ \t]*$/gi;
	let regexResult = memberStartRegex.exec(line);

	if(regexResult == null || regexResult.length < 3)
		return null;

	let visibility = regexResult[1];
	let name = regexResult[2];
	let intendention = GetNumberOfFrontSpaces(line);
	let nameStartIndex = line.indexOf(line);

	let range: ls.Range = ls.Range.create(
		statement.GetPostitionByCharacter(intendention), 
		statement.GetPostitionByCharacter(intendention + regexResult[0].trim().length)
	);
	
	let symbol: VBSMemberSymbol = new VBSMemberSymbol();
	symbol.visibility = visibility;
	symbol.type = "";
	symbol.name = name;
	symbol.args = "";
	symbol.symbolRange = range;
	symbol.nameLocation = ls.Location.create(uri, 
		ls.Range.create(
			statement.GetPostitionByCharacter(nameStartIndex),
			statement.GetPostitionByCharacter(nameStartIndex + name.length)
		)
	);
	symbol.parentName = openClassName;

	return symbol;
}

function GetVariableNamesFromList(vars: string): string[] {
	return vars.split(',').map(function(s) { return s.trim(); });
}

function GetVariableSymbol(statement: MultiLineStatement, uri: string) : VBSVariableSymbol[] {
	let line: string = statement.GetFullStatement();

	let variableSymbols: VBSVariableSymbol[] = [];
	let memberStartRegex:RegExp = /^[ \t]*(dim[ \t]+)(([a-zA-Z0-9\-\_]+[ \t]*\,[ \t]*)*)([a-zA-Z0-9\-\_]+)[ \t]*$/gi;
	let regexResult = memberStartRegex.exec(line);

	if(regexResult == null || regexResult.length < 3)
		return null;

	// (dim[ \t]+)
	let visibility = regexResult[1];
	let variables = GetVariableNamesFromList(regexResult[2] + regexResult[4]);
	let intendention = GetNumberOfFrontSpaces(line);
	let nameStartIndex = line.indexOf(line);
	let firstElementOffset = visibility.length;
	let parentName: string = "";

	if(openClassName != null)
		parentName = openClassName;

	if(openMethod != null)
		parentName = openMethod.name;

	if(openProperty != null)
		parentName = openProperty.name;

	for (let i = 0; i < variables.length; i++) {
		let varName = variables[i];
		let symbol: VBSVariableSymbol = new VBSVariableSymbol();
		symbol.visibility = "";
		symbol.type = "";
		symbol.name = varName;
		symbol.args = "";
		symbol.nameLocation = ls.Location.create(uri, 
			GetNameRange(statement, varName )
		);
		
		symbol.symbolRange = ls.Range.create(
			ls.Position.create(symbol.nameLocation.range.start.line, symbol.nameLocation.range.start.character - firstElementOffset), 
			ls.Position.create(symbol.nameLocation.range.end.line, symbol.nameLocation.range.end.character)
		);
		
		firstElementOffset = 0;
		symbol.parentName = parentName;
		
		variableSymbols.push(symbol);
	}

	return variableSymbols;
}

function GetNameRange(statement: MultiLineStatement, name: string): ls.Range {
	let line: string = statement.GetFullStatement();

	let findVariableName = new RegExp("(" + name.trim() + "[ \t]*)(\,|$)","gi");
	let matches = findVariableName.exec(line);

	let rng = ls.Range.create(
		statement.GetPostitionByCharacter(matches.index),
		statement.GetPostitionByCharacter(matches.index + name.trim().length)
	);

	return rng;
}

function GetConstantSymbol(statement: MultiLineStatement, uri: string) : VBSConstantSymbol {
	if(openMethod != null || openProperty != null)
		return null;

	let line: string = statement.GetFullStatement();

	let memberStartRegex:RegExp = /^[ \t]*(public[ \t]+|private[ \t]+)?const[ \t]+([a-zA-Z0-9\-\_]+)[ \t]*\=.*$/gi;
	let regexResult = memberStartRegex.exec(line);

	if(regexResult == null || regexResult.length < 3)
		return null;

	let visibility = regexResult[1];
	if(visibility != null)
		visibility = visibility.trim();

	let name = regexResult[2].trim();
	let intendention = GetNumberOfFrontSpaces(line);
	let nameStartIndex = line.indexOf(line);

	let range: ls.Range = ls.Range.create(
		statement.GetPostitionByCharacter(intendention), 
		statement.GetPostitionByCharacter(intendention + regexResult[0].trim().length)
	);
	
	let parentName: string = "";
	
	if(openClassName != null)
		parentName = openClassName;

	if(openMethod != null)
		parentName = openMethod.name;

	if(openProperty != null)
		parentName = openProperty.name;

	let symbol: VBSConstantSymbol = new VBSConstantSymbol();
	symbol.visibility = visibility;
	symbol.type = "";
	symbol.name = name;
	symbol.args = "";
	symbol.symbolRange = range;
	symbol.nameLocation = ls.Location.create(uri, 
		ls.Range.create(
			statement.GetPostitionByCharacter(nameStartIndex),
			statement.GetPostitionByCharacter(nameStartIndex + name.length)
		)
	);
	symbol.parentName = parentName;

	return symbol;
}

function GetClassStart(statement: MultiLineStatement, uri: string) : boolean {
	let line: string = statement.GetFullStatement();

	let classStartRegex:RegExp = /^[ \t]*class[ \t]+([a-zA-Z0-9\-\_]+)[ \t]*$/gi;
	let regexResult = classStartRegex.exec(line);

	if(regexResult == null || regexResult.length < 2)
		return false;

	let name = regexResult[1];
	openClassName = name;
	openClassStart = statement.GetPostitionByCharacter(GetNumberOfFrontSpaces(line));

	return true;
}

function GetClassSymbol(statement: MultiLineStatement, uri: string) : VBSClassSymbol {
	let line: string = statement.GetFullStatement();

	let classEndRegex:RegExp = /^[ \t]*end[ \t]+class[ \t]*$/gi;

	if(openClassName == null)
		return null;
	
	let regexResult = classEndRegex.exec(line);

	if(regexResult == null || regexResult.length < 1)
		return null;

	if(openMethod != null) {
		// ERROR! expected to close method before!
		console.error("ERROR - line " + statement.startLine + " at " + statement.startCharacter + ": 'end " + openMethod.type + "' expected!");
	}

	if(openProperty != null) {
		// ERROR! expected to close property before!
		console.error("ERROR - line " + statement.startLine + " at " + statement.startCharacter + ": 'end property' expected!");
	}

	let range: ls.Range = ls.Range.create(openClassStart, statement.GetPostitionByCharacter(regexResult[0].length))
	let symbol: VBSClassSymbol = new VBSClassSymbol();
	symbol.name = openClassName;
	symbol.nameLocation = ls.Location.create(uri, 
		ls.Range.create(openClassStart, 
			ls.Position.create(openClassStart.line, openClassStart.character + openClassName.length)
		)
	);
	symbol.symbolRange = range;
	//let symbol: ls.SymbolInformation = ls.SymbolInformation.create(openClassName, ls.SymbolKind.Class, range, uri);

	openClassName = null;
	openClassStart = ls.Position.create(-1, -1);

	return symbol;
}

// Listen on the connection
connection.listen();