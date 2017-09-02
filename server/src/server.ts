/* --------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License. See License.txt in the project root for license information.
 * ------------------------------------------------------------------------------------------ */
'use strict';

import {
	IPCMessageReader, IPCMessageWriter,
	createConnection, IConnection, TextDocumentSyncKind,
	TextDocuments, TextDocument, Diagnostic, DiagnosticSeverity,
	InitializeParams, InitializeResult, TextDocumentPositionParams,
	CompletionItem, CompletionItemKind
} from 'vscode-languageserver';

import {
	DocumentSymbolParams, SymbolInformation, SymbolKind, Range,
	Position
} from 'vscode-languageserver-types';

// Create a connection for the server. The connection uses Node's IPC as a transport
let connection: IConnection = createConnection(new IPCMessageReader(process), new IPCMessageWriter(process));

// Create a simple text document manager. The text document manager
// supports full document sync only
let documents: TextDocuments = new TextDocuments();
// Make the text document manager listen on the connection
// for open, change and close text document events
documents.listen(connection);

// After the server has started the client sends an initialize request. The server receives
// in the passed params the rootPath of the workspace plus the client capabilities. 
let workspaceRoot: string;
connection.onInitialize((params): InitializeResult => {
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
documents.onDidChangeContent((change) => {
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
connection.onDidChangeConfiguration((change) => {
	let settings = <Settings>change.settings;
	maxNumberOfProblems = settings.vbsLanguageServer.maxNumberOfProblems || 100;
});

connection.onDidChangeWatchedFiles((change) => {
});

let t: Thenable<string>;

connection.onDocumentSymbol((docParams: DocumentSymbolParams): SymbolInformation[] => {
	let symbolsList: SymbolInformation[] = [];

	CollectSymbols(documents.get(docParams.textDocument.uri), symbolsList);
	
	return symbolsList;
});

function CollectSymbols(document: TextDocument, symbols: SymbolInformation[]): void {
	let lines = document.getText().split(/\r?\n/g);
	let problems = 0;

	for (var i = 0; i < lines.length && problems < maxNumberOfProblems; i++) {
		let line = lines[i];
		let newSym = FindSymbol(line, i, document.uri);
		
		if(newSym != null)
		{
			symbols.push(newSym);
		}
	}
}

function FindSymbol(line: string, lineNumber: number, uri: string) : SymbolInformation {
	let functionStartRegex:RegExp = /^ *(public +|private +)?function +([a-zA-Z0-9\-\_]+) *(\(([a-zA-Z0-9\_\-, ]*)\))?/gi;
	let subStartRegex:RegExp = /^ *(public +|private +)?sub +([a-zA-Z0-9\-\_]+) *(\(([a-zA-Z0-9\_\-, ]*)\))?/gi;

	let newSym;
	newSym = GetSimpleSymbol(functionStartRegex, line, lineNumber, uri);
	if(newSym != null)
		return newSym;

	newSym = GetSimpleSymbol(subStartRegex, line, lineNumber, uri);
	if(newSym != null)
		return newSym;
	
	newSym = GetPropertySymbol(line, lineNumber, uri);
	if(newSym != null)
		return newSym;

	if(IsStartClass(line, lineNumber, uri))
		return null;

	newSym = GetClassSymbol(line, lineNumber, uri);
	if(newSym != null)
		return newSym;

	newSym = GetMemberSymbol(line, lineNumber, uri);
	if(newSym != null)
		return newSym;
}

let openClassName : string = null;
let openClassStart : Position = Position.create(-1, -1);

function GetSimpleSymbol(rex: RegExp, line: string, lineNumber: number, uri: string) : SymbolInformation {
	let regexResult = rex.exec(line);
	
	if(regexResult == null || regexResult.length < 5)
		return null;
	
	let visibility = regexResult[1];
	let name = regexResult[2];
	let functionArgs = regexResult[4];
	
	let range: Range = Range.create(Position.create(lineNumber, 0), Position.create(lineNumber, regexResult[0].length))
	let symbol: SymbolInformation = SymbolInformation.create(name + " (" + (functionArgs || "") + ")", SymbolKind.Function, range, uri, openClassName);
	return symbol;
}

function GetPropertySymbol(line: string, lineNumber: number, uri: string) : SymbolInformation {
	let propertyStartRegex:RegExp = /^ *(public +|private +)?property +(let +|set +|get +)([a-zA-Z0-9\-\_]+) *(\(([a-zA-Z0-9\_\-, ]*)\))?/gi;
	let regexResult = propertyStartRegex.exec(line);
	
	if(regexResult == null || regexResult.length < 4)
		return null;
	
	let visibility = regexResult[1];
	let propertyKind = regexResult[2];
	let name = regexResult[3];
	let propertyArgs = regexResult[4];
	
	let range: Range = Range.create(Position.create(lineNumber, 0), Position.create(lineNumber, regexResult[0].length))
	let symbol: SymbolInformation = SymbolInformation.create(propertyKind + " " + name + " " + (propertyArgs || ""), SymbolKind.Property, range, uri, openClassName);
	return symbol;
}

function GetMemberSymbol(line: string, lineNumber: number, uri: string) : SymbolInformation {
	let memberStartRegex:RegExp = /^ *(public +|private +)([a-zA-Z0-9\-\_]+) *($|\:|\')/gi;
	let regexResult = memberStartRegex.exec(line);
	
	if(regexResult == null || regexResult.length < 3)
		return null;
	
	let visibility = regexResult[1];
	let name = regexResult[2];
	
	let range: Range = Range.create(Position.create(lineNumber, 0), Position.create(lineNumber, regexResult[0].length))
	let symbol: SymbolInformation = SymbolInformation.create(name, SymbolKind.Field, range, uri, openClassName);
	return symbol;
}

function IsStartClass(line: string, lineNumber: number, uri: string) : boolean {
	let classStartRegex:RegExp = /^ *class +([a-zA-Z0-9\-\_]+)/gi;
	let regexResult = classStartRegex.exec(line);
	
	if(regexResult == null || regexResult.length < 2)
		return false;
	
	let name = regexResult[1];
	openClassName = name;
	openClassStart = Position.create(lineNumber, 0);
	
	return true;
}

function GetClassSymbol(line: string, lineNumber: number, uri: string) : SymbolInformation {
	let classEndRegex:RegExp = /^ *end +class( |$)/gi;
	if(openClassName == null)
		return null;

	let regexResult = classEndRegex.exec(line);
	
	if(regexResult == null || regexResult.length < 1)
		return null;

	let range: Range = Range.create(openClassStart, Position.create(lineNumber, regexResult[0].length))
	let symbol: SymbolInformation = SymbolInformation.create(openClassName, SymbolKind.Class, range, uri);

	openClassName = null;
	openClassStart = Position.create(-1, -1);

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