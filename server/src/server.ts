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
	let ci: ls.CompletionItem = ls.CompletionItem.create("hello world!");
	let scopeSymbols = GetSymbolsOfScope(symbols, textDocumentPosition.position);
	return VBSSymbol.GetLanguageServerCompletionItems(scopeSymbols);
}

function GetSymbolsOfScope(symbols: VBSSymbol[], position: ls.Position): VBSSymbol[] {
	// sort by start positition
	let sortedSymbols: VBSSymbol[] = symbols.sort(function(a: VBSSymbol, b: VBSSymbol){
		let diff = a.symbolRange.start.line - b.symbolRange.start.line;
		
		if(diff != 0)
			return diff;

		return a.symbolRange.start.character - b.symbolRange.start.character;
	});

	// bacause of hoisting we will have just a few possible scopes:
	// - file wide
	// - method of file wide
	// - class scope
	// - method or property of class scope
	
	// find out in which scope we are
	// get all symbols which are accessable from there (ignore visibility in the first step)

	// very first shot: ignore the scopes completly!
	return sortedSymbols;
}

let symbolCache: { [id: string] : VBSSymbol[]; } = {};
function RefreshDocumentsSymbols(uri: string) {
	let symbolsList: VBSSymbol[] = [];
	CollectSymbols(documents.get(uri), symbolsList);
	symbolCache[uri] = symbolsList;
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
			let newSym = FindSymbol(statement, i, document.uri);
			if(newSym != null)
			{
				symbols.push(newSym);
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

function FindSymbol(statement: string, lineNumber: number, uri: string) : VBSSymbol {
	let newSym: VBSSymbol;

	if(GetMethodStart(statement, lineNumber, uri))
		return null;

	newSym = GetMethodSymbol(statement, lineNumber, uri);
	if(newSym != null)
		return newSym;

	if(GetPropertyStart(statement, lineNumber, uri))
		return null;

	newSym = GetPropertySymbol(statement, lineNumber, uri);;
	if(newSym != null)
		return newSym;

	if(GetClassStart(statement, lineNumber, uri))
		return null;

	newSym = GetClassSymbol(statement, lineNumber, uri);
	if(newSym != null)
		return newSym;

	newSym = GetMemberSymbol(statement, lineNumber, uri);
	if(newSym != null)
		return newSym;

	newSym = GetVariableSymbol(statement, lineNumber, uri);
	if(newSym != null)
		return newSym;

	newSym = GetConstantSymbol(statement, lineNumber, uri);
	if(newSym != null)
		return newSym;
}

let openClassName : string = null;
let openClassStart : ls.Position = ls.Position.create(-1, -1);

class OpenMethod {
	visibility: string;
	type: string;
	name: string;
	args: string;
	startPosition: ls.Position;
	nameLocation: ls.Location;
}

let openMethod: OpenMethod = null;

function GetMethodStart(line: string, lineNumber: number, uri: string): boolean {
	let rex:RegExp = /^[ \t]*(public[ \t]+|private[ \t]+)?(function|sub)[ \t]+([a-zA-Z0-9\-\_]+)[ \t]*(\(([a-zA-Z0-9\_\-, \t]*)\))?[ \t]*$/gi;
	let regexResult = rex.exec(line);

	if(regexResult == null || regexResult.length < 6)
		return;

	if(openMethod == null) {
		openMethod = {
			visibility: regexResult[1],
			type: regexResult[2],
			name: regexResult[3],
			args: regexResult[5],
			startPosition: ls.Position.create(lineNumber, GetNumberOfFrontSpaces(line)),
			nameLocation: ls.Location.create(uri, ls.Range.create(
				ls.Position.create(lineNumber, line.indexOf(regexResult[3])),
				ls.Position.create(lineNumber, line.indexOf(regexResult[3]) + regexResult[3].length)))
		};
		return true;
	} else {
		// ERROR!!! I expected "end function|sub"!
	}

	return false;
}

function GetMethodSymbol(line: string, lineNumber: number, uri: string) : VBSMethodSymbol{
	let classEndRegex:RegExp = /^[ \t]*end[ \t]+(function|sub)[ \t]*$/gi;

	let regexResult = classEndRegex.exec(line);

	if(regexResult == null || regexResult.length < 2)
		return null;

	let type = regexResult[1];

	if(openMethod == null) {
		// ERROR!!! I cannot close any method!
		return null;
	}

	if(type != openMethod.type) {
		// ERROR!!! I expected end function|sub and not sub|function!
		// show the user the error and then go on like it was the right type!
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

	openMethod = null;

	return symbol;
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
	args: string;
	startPosition: ls.Position;
	nameLocation: ls.Location;
}

let openProperty: OpenProperty = null;

function GetPropertyStart(line: string, lineNumber: number, uri: string) : boolean {
	let propertyStartRegex:RegExp = /^[ \t]*(public[ \t]+|private[ \t]+)?property[ \t]+(let[ \t]+|set[ \t]+|get[ \t]+)([a-zA-Z0-9\-\_]+)[ \t]*(\(([a-zA-Z0-9\_\-, \t]*)\))?[ \t]*$/gi;
	let regexResult = propertyStartRegex.exec(line);

	if(regexResult == null || regexResult.length < 6)
		return null;

	if(openProperty == null) {
		openProperty = {
			visibility: regexResult[1],
			type: regexResult[2],
			name: regexResult[3],
			args: regexResult[5],
			startPosition: ls.Position.create(lineNumber, GetNumberOfFrontSpaces(line)),
			nameLocation: ls.Location.create(uri, ls.Range.create(
				ls.Position.create(lineNumber, line.indexOf(regexResult[3])),
				ls.Position.create(lineNumber, line.indexOf(regexResult[3]) + regexResult[3].length)))
		};

		return true;
	} else {
		// ERROR!!! I expected "end function|sub"!
	}

	return false;
}

function GetPropertySymbol(statement: string, lineNumber: number, uri: string) : VBSPropertySymbol{
	let classEndRegex:RegExp = /^[ \t]*end[ \t]+property[ \t]*$/gi;

	let regexResult = classEndRegex.exec(statement);

	if(regexResult == null || regexResult.length < 1)
		return null;

	if(openProperty == null) {
		// ERROR!!! I cannot close any method!
		return null;
	}

	// range of the whole definition
	let range: ls.Range = ls.Range.create(openProperty.startPosition, ls.Position.create(lineNumber, GetNumberOfFrontSpaces(statement) + regexResult[0].trim().length))
	
	let symbol = new VBSPropertySymbol()
	symbol.type = openProperty.type;
	symbol.name = openProperty.name;
	symbol.args = openProperty.args;
	symbol.symbolRange = range;
	symbol.nameLocation = openProperty.nameLocation;
	symbol.parentName = openClassName;
	symbol.symbolRange = range;

	openProperty = null;

	return symbol;
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

function GetVariableSymbol(line: string, lineNumber: number, uri: string) : VBSVariableSymbol {
	let memberStartRegex:RegExp = /^[ \t]*(dim[ \t]+)([a-zA-Z0-9\-\_]+)[ \t]*$/gi;
	let regexResult = memberStartRegex.exec(line);

	if(regexResult == null || regexResult.length < 3)
		return null;

	// (dim[ \t]+)
	let visibility = regexResult[1];
	let name = regexResult[2];
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

	let symbol: VBSVariableSymbol = new VBSVariableSymbol();
	symbol.visibility = "";
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

	if(openMethod != null) {
		// ERROR! expected to close method before!
	}

	if(openProperty != null) {
		// ERROR! expected to close property before!
	}

	let regexResult = classEndRegex.exec(line);

	if(regexResult == null || regexResult.length < 1)
		return null;

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