/* --------------------------------------------------------------------------------------------
 * Copyright (c) Andreas Lenzen. All rights reserved.
 * Licensed under the MIT License. See License.txt in the project root for license information.
 * ------------------------------------------------------------------------------------------ */
'use strict';

import * as ls from 'vscode-languageserver';

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
			documentSymbolProvider: true
		}
	}
});

// The content of a text document has changed. This event is emitted
// when the text document first opened or when its content has changed.
documents.onDidChangeContent((change: ls.TextDocumentChangeEvent) => {
});

// The settings interface describe the server relevant settings part
interface Settings {
	vbsLanguageServer: ExampleSettings;
}

// These are the example settings we defined in the client's package.json
// file
interface ExampleSettings {
	maxNumberOfProblems: number;
}

// hold the maxNumberOfProblems setting
let maxNumberOfProblems: number;
// The settings have changed. Is send on server activation
// as well.
connection.onDidChangeConfiguration((change: ls.DidChangeConfigurationParams) => {
	let settings = <Settings>change.settings;
	maxNumberOfProblems = settings.vbsLanguageServer.maxNumberOfProblems || 100;
});

connection.onDidChangeWatchedFiles((changeParams: ls.DidChangeWatchedFilesParams) => {
	for (var i = 0; i < changeParams.changes.length; i++) {
		var event = changeParams.changes[i];
		
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

let symbolCache: { [id: string] : ls.SymbolInformation[]; } = {}; 
function RefreshDocumentsSymbols(uri: string) {
	let symbolsList: ls.SymbolInformation[] = [];
	CollectSymbols(documents.get(uri), symbolsList);
	symbolCache[uri] = symbolsList;
}

function GetSymbolsOfDocument(uri: string) : ls.SymbolInformation[] {
	RefreshDocumentsSymbols(uri);
	return symbolCache[uri];
}

function GetWorkspaceSymbols(query: string) : ls.SymbolInformation[] {
	let symbolsList: ls.SymbolInformation[] = [];

	for(let key in symbolCache) {
		for (var i = 0; i < symbolCache[key].length; i++) {
			var symbol = symbolCache[key][i];
			if(SymbolMatchesQuery(symbol, query))
				symbolsList.push(symbol);
		}
	}
	
	return symbolsList;
}

function SymbolMatchesQuery(symbol: ls.SymbolInformation, query: string): boolean {
	return symbol.name.indexOf(query) > -1;
}

let t: Thenable<string>;

connection.onDocumentSymbol((docParams: ls.DocumentSymbolParams): ls.SymbolInformation[] => {
	return GetSymbolsOfDocument(docParams.textDocument.uri);
});

function CollectSymbols(document: ls.TextDocument, symbols: ls.SymbolInformation[]): void {
	let lines = document.getText().split(/\r?\n/g);
	let problems = 0;

	for (var i = 0; i < lines.length && problems < maxNumberOfProblems; i++) {
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

function FindSymbol(statement: string, lineNumber: number, uri: string) : ls.SymbolInformation {
	let newSym;
	
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
}

let openClassName : string = null;
let openClassStart : ls.Position = ls.Position.create(-1, -1);

interface IOpenMethod {
	visibility: string;
	type: string;
	name: string;
	args: string;
	startPosition: ls.Position;
	nameLocation: ls.Location;
}

let openMethod: IOpenMethod = null;

function GetMethodStart(line: string, lineNumber: number, uri: string): boolean {
	let rex:RegExp = /^ *(public +|private +)?(function|sub) +([a-zA-Z0-9\-\_]+) *(\(([a-zA-Z0-9\_\-, ]*)\))? *$/gi;
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

function GetMethodSymbol(line: string, lineNumber: number, uri: string) : ls.SymbolInformation{
	let classEndRegex:RegExp = /^ *end +(function|sub) *$/gi;

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
	let symbol: ls.SymbolInformation = ls.SymbolInformation.create(
		openMethod.name + " (" + (openMethod.args || "") + ")", 
		(openClassName == null ? ls.SymbolKind.Function : ls.SymbolKind.Method), 
		range,
		uri,
		openClassName);
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

interface IOpenProperty {
	visibility: string;
	type: string;
	name: string;
	args: string;
	startPosition: ls.Position;
	nameLocation: ls.Location;
}

let openProperty: IOpenProperty = null;

function GetPropertyStart(line: string, lineNumber: number, uri: string) : boolean {
	let propertyStartRegex:RegExp = /^ *(public +|private +)?property +(let +|set +|get +)([a-zA-Z0-9\-\_]+) *(\(([a-zA-Z0-9\_\-, ]*)\))? *$/gi;
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

function GetPropertySymbol(statement: string, lineNumber: number, uri: string) : ls.SymbolInformation{
	let classEndRegex:RegExp = /^ *end +property *$/gi;
	
	let regexResult = classEndRegex.exec(statement);
	
	if(regexResult == null || regexResult.length < 1)
		return null;
	
	if(openProperty == null) {
		// ERROR!!! I cannot close any method!
		return null;
	}
	
	// range of the whole definition
	let range: ls.Range = ls.Range.create(openProperty.startPosition, ls.Position.create(lineNumber, GetNumberOfFrontSpaces(statement) + regexResult[0].trim().length))
	let symbol: ls.SymbolInformation = ls.SymbolInformation.create(
		openProperty.type + "" +  openProperty.name + " (" + (openProperty.args || "") + ")", 
		ls.SymbolKind.Property, 
		range,
		uri,
		openClassName);

	openProperty = null;

	return symbol;
}

function GetMemberSymbol(line: string, lineNumber: number, uri: string) : ls.SymbolInformation {
	let memberStartRegex:RegExp = /^ *(public +|private +)([a-zA-Z0-9\-\_]+) *$/gi;
	let regexResult = memberStartRegex.exec(line);
	
	if(regexResult == null || regexResult.length < 3)
		return null;
	
	let visibility = regexResult[1];
	let name = regexResult[2];
	let intendention = GetNumberOfFrontSpaces(line);

	let range: ls.Range = ls.Range.create(ls.Position.create(lineNumber, intendention), ls.Position.create(lineNumber, intendention + regexResult[0].trim().length))
	let symbol: ls.SymbolInformation = ls.SymbolInformation.create(name, ls.SymbolKind.Field, range, uri, openClassName);
	return symbol;
}

function GetVariableSymbol(line: string, lineNumber: number, uri: string) : ls.SymbolInformation {
	if(openClassName != null || openMethod != null || openProperty != null)
		return null;

	let memberStartRegex:RegExp = /^ *(dim +)([a-zA-Z0-9\-\_]+) *$/gi;
	let regexResult = memberStartRegex.exec(line);
	
	if(regexResult == null || regexResult.length < 3)
		return null;
	
	let visibility = regexResult[1];
	let name = regexResult[2];
	let intendention = GetNumberOfFrontSpaces(line);

	let range: ls.Range = ls.Range.create(ls.Position.create(lineNumber, intendention), ls.Position.create(lineNumber, intendention + regexResult[0].trim().length))
	let symbol: ls.SymbolInformation = ls.SymbolInformation.create(name, ls.SymbolKind.Variable, range, uri, null);
	return symbol;
}

function GetClassStart(line: string, lineNumber: number, uri: string) : boolean {
	let classStartRegex:RegExp = /^ *class +([a-zA-Z0-9\-\_]+) *$/gi;
	let regexResult = classStartRegex.exec(line);
	
	if(regexResult == null || regexResult.length < 2)
		return false;
	
	let name = regexResult[1];
	openClassName = name;
	openClassStart = ls.Position.create(lineNumber, 0);
	
	return true;
}

function GetClassSymbol(line: string, lineNumber: number, uri: string) : ls.SymbolInformation {
	let classEndRegex:RegExp = /^ *end +class *$/gi;
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
	let symbol: ls.SymbolInformation = ls.SymbolInformation.create(openClassName, ls.SymbolKind.Class, range, uri);

	openClassName = null;
	openClassStart = ls.Position.create(-1, -1);

	return symbol;
}

/*
connection.onDidOpenTextDocument((params) => {
	// A text document got opened in VSCode.
	// params.textDocument.uri uniquely identifies the document. For documents store on disk this is a file URI.
	// params.textDocument.text the initial full content of the document.
	connection.console.log(`${params.textDocument.uri} opened.`);
});

connection.onDidChangeTextDocument((params) => {
	// The content of a text document did change in VSCode.
	// params.textDocument.uri uniquely identifies the document.
	// params.contentChanges describe the content changes to the document.
	connection.console.log(`${params.textDocument.uri} changed: ${JSON.stringify(params.contentChanges)}`);
});

connection.onDidCloseTextDocument((params) => {
	// A text document got closed in VSCode.
	// params.textDocument.uri uniquely identifies the document.
	connection.console.log(`${params.textDocument.uri} closed.`);
});
*/

// Listen on the connection
connection.listen();