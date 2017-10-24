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
	connection.console.log("Start refreshing symbols...");
	let symbolsList: VBSSymbol[] = [];
	CollectSymbols(documents.get(uri), symbolsList);
	symbolCache[uri] = symbolsList;
	connection.console.log("Symbols refreshed!");
}

connection.onDocumentSymbol((docParams: ls.DocumentSymbolParams): ls.SymbolInformation[] => {
	return GetSymbolsOfDocument(docParams.textDocument.uri);
});

function CollectSymbols(document: ls.TextDocument, symbols: VBSSymbol[]): void {
	let lines = document.getText().split(/\r?\n/g);

	for (var i = 0; i < lines.length; i++) {
		let line = lines[i];

		let containsComment = line.indexOf("'");
		if(containsComment > -1)
			line = line.substring(0, containsComment);

		line = ReplaceStringLiterals(line);
		let statements = SplitStatements(line);

		statements.forEach(statement => {
			let newSymbols = FindSymbol(statement, i, document.uri);

			for (var j = 0; j < newSymbols.length; j++) {
				var newSym = newSymbols[j];
				
				if(newSym != null) {
					symbols.push(newSym);
				}
			}
		});
	}
}

function SplitStatements(line: string) {
	let statements: string[] = line.split(":");
	let offset = 0;

	for (var i = 0; i < statements.length; i++) {
		statements[i] = " ".repeat(offset) + statements[i];
		offset = statements[i].length + 1;
	}

	return statements;
}

function ReplaceStringLiterals(line:string) : string {
	let stringLiterals = /\"(([^\"]|\"\")*)\"/gi;
	return line.replace(stringLiterals, ReplaceBySpaces);
}

function ReplaceBySpaces(match: string) : string {
	return " ".repeat(match.length);
}

function FindSymbol(statement: string, lineNumber: number, uri: string) : VBSSymbol[] {
	let newSym: VBSSymbol;
	let newSyms: VBSVariableSymbol[] = null;

	if(GetMethodStart(statement, lineNumber, uri))
		return [];

	newSyms = GetMethodSymbol(statement, lineNumber, uri);
	if(newSyms != null && newSyms.length != 0)
		return newSyms;

	if(GetPropertyStart(statement, lineNumber, uri))
		return [];

	newSyms = GetPropertySymbol(statement, lineNumber, uri);;
	if(newSyms != null && newSyms.length != 0)
		return newSyms;

	if(GetClassStart(statement, lineNumber, uri))
		return [];

	newSym = GetClassSymbol(statement, lineNumber, uri);
	if(newSym != null)
		return [newSym];

	newSym = GetMemberSymbol(statement, lineNumber, uri);
	if(newSym != null)
		return [newSym];

	newSyms = GetVariableSymbol(statement, lineNumber, uri);
	if(newSyms != null && newSyms.length != 0)
		return newSyms;

	newSym = GetConstantSymbol(statement, lineNumber, uri);
	if(newSym != null)
		return [newSym];

	return [];
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
}

let openMethod: OpenMethod = null;

function GetMethodStart(line: string, lineNumber: number, uri: string): boolean {
	let rex:RegExp = /^[ \t]*(public[ \t]+|private[ \t]+)?(function|sub)([ \t]+)([a-zA-Z0-9\-\_]+)([ \t]*)(\(([a-zA-Z0-9\_\-, \t]*)\))?[ \t]*$/gi;
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
			startPosition: ls.Position.create(lineNumber, leadingSpaces),
			nameLocation: ls.Location.create(uri, ls.Range.create(
				ls.Position.create(lineNumber, line.indexOf(regexResult[3])),
				ls.Position.create(lineNumber, line.indexOf(regexResult[3]) + regexResult[3].length)))
		};
		
		if(openMethod.args == null)
			openMethod.args = "";

		return true;
	} else {
		// ERROR!!! I expected "end function|sub"!
		console.log("ERROR - line " + lineNumber + ": 'end function' or 'end sub' expected!");
	}

	return false;
}

function GetMethodSymbol(line: string, lineNumber: number, uri: string) : VBSSymbol[] {
	let classEndRegex:RegExp = /^[ \t]*end[ \t]+(function|sub)[ \t]*$/gi;

	let regexResult = classEndRegex.exec(line);

	if(regexResult == null || regexResult.length < 2)
		return null;

	let type = regexResult[1];

	if(openMethod == null) {
		// ERROR!!! I cannot close any method!
		console.log("ERROR - line " + lineNumber + ": There is no " + type + " to end!");
		return null;
	}

	if(type != openMethod.type) {
		// ERROR!!! I expected end function|sub and not sub|function!
		// show the user the error and then go on like it was the right type!
		console.log("ERROR - line " + lineNumber + ": 'end " + openMethod.type + "' expected!");
	}

	let range: ls.Range = ls.Range.create(openMethod.startPosition, ls.Position.create(lineNumber, GetNumberOfFrontSpaces(line) + regexResult[0].trim().length))
	
	let symbol: VBSMethodSymbol = new VBSMethodSymbol();
	symbol.visibility = openMethod.visibility;
	symbol.type = openMethod.type;
	symbol.name = openMethod.name;
	symbol.args = openMethod.args;
	symbol.nameLocation = openMethod.nameLocation;
	symbol.parentName = openClassName;
	symbol.symbolRange = range;

	let parametersSymbol = GetParameterSymbols(openMethod.args, openMethod.argsIndex, range.start.line, uri);

	openMethod = null;

	//return [symbol];
	return parametersSymbol.concat(symbol);
}

function ReplaceAll(target: string, search: string, replacement: string): string {
    return target.replace(new RegExp(search, 'g'), replacement);
};

function GetParameterSymbols(args: string, argsIndex: number, lineNumber: number, uri: string): VBSVariableSymbol[] {
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
			ls.Position.create(lineNumber, argsIndex + arg.indexOf(varSymbol.name)),
			ls.Position.create(lineNumber, argsIndex + arg.indexOf(varSymbol.name) + varSymbol.name.length)
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
}

let openProperty: OpenProperty = null;

function GetPropertyStart(line: string, lineNumber: number, uri: string) : boolean {
	let propertyStartRegex:RegExp = /^[ \t]*(public[ \t]+|private[ \t]+)?(property[ \t]+)(let[ \t]+|set[ \t]+|get[ \t]+)([a-zA-Z0-9\-\_]+)([ \t]*)(\(([a-zA-Z0-9\_\-, \t]*)\))?[ \t]*$/gi;
	let regexResult = propertyStartRegex.exec(line);

	if(regexResult == null || regexResult.length < 6)
		return null;

	let leadingSpaces = GetNumberOfFrontSpaces(line);
	let preLength = leadingSpaces + regexResult.index;
	
	for (var i = 1; i < 6; i++) {
		var resElement = regexResult[i];
		if(resElement != null)
			preLength += resElement.length;
	}

	if(openProperty == null) {
		openProperty = {
			visibility: regexResult[1],
			type: regexResult[3],
			name: regexResult[4],
			argsIndex: preLength + 1,
			args: regexResult[7],
			startPosition: ls.Position.create(lineNumber, leadingSpaces),
			nameLocation: ls.Location.create(uri, ls.Range.create(
				ls.Position.create(lineNumber, line.indexOf(regexResult[4])),
				ls.Position.create(lineNumber, line.indexOf(regexResult[4]) + regexResult[4].length)))
		};

		if(openProperty.args == null)
			openProperty.args = "";

		return true;
	} else {
		// ERROR!!! I expected "end function|sub"!
		console.log("ERROR - line " + lineNumber + ": 'end function' or 'end sub' expected!");
	}

	return false;
}

function GetPropertySymbol(statement: string, lineNumber: number, uri: string) : VBSSymbol[] {
	let classEndRegex:RegExp = /^[ \t]*end[ \t]+property[ \t]*$/gi;

	let regexResult = classEndRegex.exec(statement);

	if(regexResult == null || regexResult.length < 1)
		return null;

	if(openProperty == null) {
		// ERROR!!! I cannot close any property!
		console.log("ERROR - line " + lineNumber + ": There is no property to end!");
		return null;
	}

	// range of the whole definition
	let range: ls.Range = ls.Range.create(openProperty.startPosition, ls.Position.create(lineNumber, GetNumberOfFrontSpaces(statement) + regexResult[0].trim().length))
	
	let symbol = new VBSPropertySymbol()
	symbol.visibility = "";
	symbol.type = openProperty.type;
	symbol.name = openProperty.name;
	symbol.args = openProperty.args;
	symbol.symbolRange = range;
	symbol.nameLocation = openProperty.nameLocation;
	symbol.parentName = openClassName;
	symbol.symbolRange = range;

	let parametersSymbol = GetParameterSymbols(openProperty.args, openProperty.argsIndex, range.start.line, uri);

	openProperty = null;

	return parametersSymbol.concat(symbol);
}

function GetMemberSymbol(line: string, lineNumber: number, uri: string) : VBSMemberSymbol {
	let memberStartRegex:RegExp = /^[ \t]*(public[ \t]+|private[ \t]+)([a-zA-Z0-9\-\_]+)[ \t]*$/gi;
	let regexResult = memberStartRegex.exec(line);

	if(regexResult == null || regexResult.length < 3)
		return null;

	let visibility = regexResult[1];
	let name = regexResult[2];
	let intendention = GetNumberOfFrontSpaces(line);
	let nameStartIndex = line.indexOf(line);

	let range: ls.Range = ls.Range.create(ls.Position.create(lineNumber, intendention), ls.Position.create(lineNumber, intendention + regexResult[0].trim().length))
	
	let symbol: VBSMemberSymbol = new VBSMemberSymbol();
	symbol.visibility = visibility;
	symbol.type = "";
	symbol.name = name;
	symbol.args = "";
	symbol.symbolRange = range;
	symbol.nameLocation = ls.Location.create(uri, 
		ls.Range.create(
			ls.Position.create(lineNumber, nameStartIndex),
			ls.Position.create(lineNumber, nameStartIndex + name.length)
		)
	);
	symbol.parentName = openClassName;

	return symbol;
}

function GetVariableNamesFromList(vars: string): string[] {
	return vars.split(',').map(function(s) { return s.trim(); });
}

function GetVariableSymbol(line: string, lineNumber: number, uri: string) : VBSVariableSymbol[] {
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
			GetNameRange(lineNumber, line, varName )
		);
		symbol.symbolRange = ls.Range.create(
			ls.Position.create(lineNumber, symbol.nameLocation.range.start.character - firstElementOffset), 
			ls.Position.create(lineNumber, symbol.nameLocation.range.end.character)
		);
		firstElementOffset = 0;
		symbol.parentName = parentName;
		
		variableSymbols.push(symbol);
	}

	return variableSymbols;
}

function GetNameRange(lineNumber: number, line: string, name: string): ls.Range {
	let findVariableName = new RegExp("(" + name.trim() + "[ \t]*)(\,|$)","gi");
	let matches = findVariableName.exec(line);

	let rng = ls.Range.create(
		ls.Position.create(lineNumber, matches.index),
		ls.Position.create(lineNumber, matches.index + name.trim().length)
	)

	return rng;
}

function GetConstantSymbol(line: string, lineNumber: number, uri: string) : VBSConstantSymbol {
	if(openMethod != null || openProperty != null)
		return null;

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

	let range: ls.Range = ls.Range.create(ls.Position.create(lineNumber, intendention), ls.Position.create(lineNumber, intendention + regexResult[0].trim().length))
	
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
			ls.Position.create(lineNumber, nameStartIndex),
			ls.Position.create(lineNumber, nameStartIndex + name.length)
		)
	);
	symbol.parentName = parentName;

	return symbol;
}

function GetClassStart(line: string, lineNumber: number, uri: string) : boolean {
	let classStartRegex:RegExp = /^[ \t]*class[ \t]+([a-zA-Z0-9\-\_]+)[ \t]*$/gi;
	let regexResult = classStartRegex.exec(line);

	if(regexResult == null || regexResult.length < 2)
		return false;

	let name = regexResult[1];
	openClassName = name;
	openClassStart = ls.Position.create(lineNumber, 0);

	return true;
}

function GetClassSymbol(line: string, lineNumber: number, uri: string) : VBSClassSymbol {
	let classEndRegex:RegExp = /^[ \t]*end[ \t]+class[ \t]*$/gi;

	if(openClassName == null)
		return null;
	
	let regexResult = classEndRegex.exec(line);

	if(regexResult == null || regexResult.length < 1)
		return null;

	if(openMethod != null) {
		// ERROR! expected to close method before!
		console.log("ERROR - line " + lineNumber + ": 'end " + openMethod.type + "' expected!");
	}

	if(openProperty != null) {
		// ERROR! expected to close property before!
		console.log("ERROR - line " + lineNumber + ": 'end property' expected!");
	}

	let range: ls.Range = ls.Range.create(openClassStart, ls.Position.create(lineNumber, regexResult[0].length))
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