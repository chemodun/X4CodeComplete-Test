// The module 'vscode' contains the VS Code extensibility API
// Import the module and reference it with the alias vscode in your code below
import * as vscode from 'vscode';
import * as fs from 'fs';
import * as xml2js from 'xml2js';
import * as path from 'path';
import * as https from 'https'; // Import the built-in https module
import { JSDOM } from 'jsdom'; // Import jsdom
import TurndownService from 'turndown'; // Import TurndownService

// this method is called when your extension is activated
// your extension is activated the very first time the command is executed
let exceedinglyVerbose: boolean = false;
let limitLanguage: boolean = false;
let rootpath: string;
let extensionsFolder: string;
const languageFiles: Map<string, any> = new Map();
let luaFunctionInfo: Map<string, string> = new Map();
const luaTypeDefinitions: Map<string, string> = new Map(); // Store type definitions
const LuaFunctionCompletionItems: vscode.CompletionItem[] = [];
// Register Lua-specific features
const sel: vscode.DocumentSelector = { language: 'lua' };

// Initialize TurndownService
const turndownService = new TurndownService();

// Add settings validation function
function validateSettings(config: vscode.WorkspaceConfiguration): boolean {
  const requiredSettings = ['unpackedFileLocation', 'extensionsFolder'];

  let isValid = true;
  requiredSettings.forEach((setting) => {
    if (!config.get(setting)) {
      vscode.window.showErrorMessage(`Missing required setting: ${setting}. Please update your VSCode settings.`);
      isValid = false;
    }
  });

  return isValid;
}

// load and parse language files
function loadLanguageFiles(basePath: string, extensionsFolder: string) {
  console.log('Loading Language Files');
  try {
    // Array to store all valid 't' directory paths
    const tDirectories: string[] = [];

    // Check root 't' directory under unpackedFileLocation
    const rootTPath = path.join(basePath, 't');
    if (fs.existsSync(rootTPath) && fs.statSync(rootTPath).isDirectory()) {
      tDirectories.push(rootTPath);
    }
    // Check 't' directories under languageFilesFolder subdirectories
    if (fs.existsSync(extensionsFolder) && fs.statSync(extensionsFolder).isDirectory()) {
      const subdirectories = fs
        .readdirSync(extensionsFolder, { withFileTypes: true })
        .filter((item) => item.isDirectory())
        .map((item) => item.name);

      for (const subdir of subdirectories) {
        const tPath = path.join(extensionsFolder, subdir, 't');
        if (fs.existsSync(tPath) && fs.statSync(tPath).isDirectory()) {
          tDirectories.push(tPath);
        }
      }
    }

    // Process all found 't' directories
    for (const tDir of tDirectories) {
      const files = fs.readdirSync(tDir);
      const languageFilesFiltered = files.filter((file) => file.startsWith('0001') && file.endsWith('.xml'));

      for (const file of languageFilesFiltered) {
        const filePath = path.join(tDir, file);
        try {
          const rawData = fs.readFileSync(filePath, 'utf-8');
          xml2js.parseString(rawData, (err: any, result: any) => {
            if (err) {
              console.error(`Error parsing ${file} in ${tDir}: ${err}`);
              return;
            }
            languageFiles.set(filePath, result);
          });
        } catch (fileError) {
          console.error(`Error reading ${file} in ${tDir}: ${fileError}`);
        }
      }
    }

    console.log(`Loaded ${languageFiles.size} language files from ${tDirectories.length} 't' directories`);
  } catch (error) {
    console.error(`Error loading language files: ${error}`);
    vscode.window.showErrorMessage('Failed to load language files from t directories');
  }
}

function findLanguageText(pageId: string, textId: string): string {
  const config = vscode.workspace.getConfiguration('x4CodeComplete-lua');
  const preferredLanguageNumber = config.get('languageNumber') || '44';
  limitLanguage = config.get('limitLanguageOutput') || false;

  interface Match {
    fileNumber: string;
    text: string;
  }
  const allMatches: Match[] = [];

  for (const [filePath, xmlData] of languageFiles) {
    // To process the diff format of the translation files
    if (!(xmlData?.language?.page || Array.isArray(xmlData?.diff?.add))) continue;

    const pages = [...(xmlData.language?.page || []), ...(xmlData.diff?.add?.flatMap((item: any) => item.page) || [])];

    const page = pages.find((p: any) => p?.$?.id === pageId);

    if (page?.t) {
      const text = page.t.find((t: any) => t?.$?.id === textId);
      if (text?._) {
        const fileName = path.basename(filePath);
        let fileNumber: string;

        if (fileName === '0001.xml') {
          fileNumber = '*'; // Special case for 0001.xml
        } else {
          // Try matching 0001-<letter><number>.xml (e.g., 0001-l007.xml)
          const matchWithLetter = fileName.match(/0001-[a-zA-Z](\d+)\.xml$/);
          if (matchWithLetter && matchWithLetter[1]) {
            fileNumber = matchWithLetter[1].replace(/^0+/, '');
          } else {
            // Original match for 0001-<number>.xml
            const match = fileName.match(/0001-(\d+)\.xml$/);
            if (match && match[1]) {
              fileNumber = match[1].replace(/^0+/, '');
            } else {
              // Fallback for any number after 0001- or 0001
              const fallbackMatch = fileName.match(/0001-?[a-zA-Z]?(\d+)/);
              fileNumber = fallbackMatch && fallbackMatch[1] ? fallbackMatch[1].replace(/^0+/, '') : 'Unknown';
            }
          }
        }

        if (!limitLanguage || fileNumber == '*' || fileNumber == preferredLanguageNumber) {
          allMatches.push({
            fileNumber,
            text: text._.split('\n')
              .map((line: string) => `${fileNumber}: ${line}`)
              .join('\n'),
          });
        }
      }
    }
  }

  allMatches.sort((a, b) => {
    if (a.fileNumber === preferredLanguageNumber && b.fileNumber !== preferredLanguageNumber) return -1;
    if (b.fileNumber === preferredLanguageNumber && a.fileNumber !== preferredLanguageNumber) return 1;
    return a.fileNumber.localeCompare(b.fileNumber);
  });

  return allMatches.length > 0 ? allMatches.map((match) => match.text).join('\n\n') : '';
}

// Key for storing parsed data in globalState
const LUA_FUNCTION_INFO_KEY = 'luaFunctionInfo';

// Save parsed data to globalState
function saveParsedDataToGlobalState(context: vscode.ExtensionContext) {
  try {
    const data = Array.from(luaFunctionInfo.entries());
    context.globalState.update(LUA_FUNCTION_INFO_KEY, data);
    console.log('Parsed data saved to globalState successfully.');
  } catch (error) {
    console.error(`Failed to save parsed data to globalState: ${error.message}`);
  }
}

// Load parsed data from globalState
function loadParsedDataFromGlobalState(context: vscode.ExtensionContext) {
  try {
    const data = context.globalState.get<[string, string][]>(LUA_FUNCTION_INFO_KEY);
    if (data) {
      luaFunctionInfo = new Map(data);
      console.log('Loaded %s Lua functions from globalState successfully.', luaFunctionInfo.size);
    } else {
      console.log('No saved Lua functions found in globalState. Fetching new data...');
    }
  } catch (error) {
    console.error(`Failed to load Lua functions from globalState: ${error.message}`);
  }
}

// Recursively find all .lua files in a directory
function findLuaFilesRecursively(directory: string): string[] {
  let luaFiles: string[] = [];
  const entries = fs.readdirSync(directory, { withFileTypes: true });

  for (const entry of entries) {
    const fullPath = path.join(directory, entry.name);
    if (entry.isDirectory()) {
      luaFiles = luaFiles.concat(findLuaFilesRecursively(fullPath));
    } else if (entry.isFile() && entry.name.endsWith('.lua')) {
      luaFiles.push(fullPath);
    }
  }

  return luaFiles;
}

// Parse a single line or block from ffi.cdef as either a type definition or a function definition
function parseFfiBlock(lines: string[]) {
  let typedefBuffer: string[] = [];
  const typeDefinitionStartPattern = /^typedef\s+(struct|enum|union)\s*\{\s*$/;
  const typeDefinitionEndPattern = /^\}\s+(\w+);$/;
  const typeDefinitionPattern = /^typedef\s+(struct|enum|union|\w+)(\s+\{.*\}|\s+)\s*(\w+);$/;
  const functionDefinitionPattern = /^(?:([\w\s*]+)\s+)?(\w+)\s*\((.*?)\);$/;

  lines.forEach((line) => {
    if (typedefBuffer.length > 0 || typeDefinitionStartPattern.test(line) || typeDefinitionPattern.test(line)) {
      // Handle multi-line typedef
      if (typedefBuffer.length > 0 || typeDefinitionStartPattern.test(line)) {
        typedefBuffer.push(line);
      }
      if (typeDefinitionEndPattern.test(line) || typeDefinitionPattern.test(line)) {
        let typedefContent = line;
        if (typedefBuffer.length > 0) {
          typedefContent = typedefBuffer.join(' ');
        }
        const typeNameMatch = typedefContent.match(typeDefinitionPattern);
        if (typeNameMatch) {
          const typeName = typeNameMatch[3];
          const typeDescription = typedefContent.replace(typeName, `\`${typeName}\``);
          luaTypeDefinitions.set(typeName, typeDescription);
        }
        typedefBuffer = []; // Clear buffer after processing
      }
      return;
    }

    // Handle function definitions
    const functionMatch = line.match(functionDefinitionPattern);
    if (functionMatch) {
      const returnType = functionMatch[1]?.trim() || 'void';
      const functionName = functionMatch[2];
      const functionNameWithPrefix = `C.${functionName}`;
      const parameters = functionMatch[3];
      const parameterPairs = parameters.split(',');
      const parametersFormatted = parameterPairs
        .map((param) => param.replace(param.trim().split(/\s+/)[1], `\`${param.trim().split(/\s+/)[1]}\``))
        .join(', ');
      let description = `${returnType} \`${functionNameWithPrefix}\`(${parametersFormatted})\n`;

      // Add type descriptions for return type
      if (luaTypeDefinitions.has(returnType)) {
        description += `\n***\n${luaTypeDefinitions.get(returnType)}`;
      }

      // Add type descriptions for parameters
      const paramTypes = parameterPairs.map((param) => param.trim().split(/\s+/)[0]);
      let firstParam = true;
      paramTypes.forEach((paramType) => {
        if (paramType === 'void') {
          return;
        } else if (paramType.endsWith('*')) {
          paramType = paramType.slice(0, -1); // Remove trailing '*'
        }
        if (luaTypeDefinitions.has(paramType)) {
          if (firstParam) {
            description += '\n';
            firstParam = false;
          }
          description += `- ${luaTypeDefinitions.get(paramType)}\n`;
        }
      });

      if (luaFunctionInfo.has(functionName)) {
        // Merge descriptions if duplicate
        const existingDescription = luaFunctionInfo.get(functionName);
        luaFunctionInfo.delete(functionName);
        luaFunctionInfo.set(functionNameWithPrefix, `${description}\n\n${existingDescription}`);
      } else if (!luaFunctionInfo.has(functionNameWithPrefix)) {
        luaFunctionInfo.set(functionNameWithPrefix, description);
      }
    }
  });
}

// Load and parse Lua files for ffi.cdef type and function definitions in a single pass
function loadLuaFunctionsInfoFromLocalFiles(basePath: string) {
  console.log('Loading Lua Definitions (types and functions) from local Lua files');
  const uiFolderPath = path.join(basePath, 'ui');
  if (!fs.existsSync(uiFolderPath) || !fs.statSync(uiFolderPath).isDirectory()) {
    console.warn(`UI folder not found at path: ${uiFolderPath}`);
    return;
  }
  const previousLuaFunctionCount = luaFunctionInfo.size;
  const luaFiles = findLuaFilesRecursively(uiFolderPath);
  luaFiles.forEach((filePath) => {
    try {
      const content = fs.readFileSync(filePath, 'utf-8');
      const ffiMatch = content.match(
        /local\s+ffi\s*=\s*require\("ffi"\)\s*local\s+C\s*=\s*ffi\.C\s*ffi\.cdef\[\[([\s\S]*?)\]\]/
      );
      if (ffiMatch && ffiMatch[1]) {
        const ffiBlock = ffiMatch[1];
        const lines = ffiBlock
          .split('\n')
          .map((line) => line.trim())
          .filter((line) => line);

        parseFfiBlock(lines);
      }
    } catch (error) {
      console.error(`Error reading or parsing Lua file: ${filePath}, Error: ${error.message}`);
    }
  });
  console.log(
    `Loaded ${luaFunctionInfo.size - previousLuaFunctionCount} Lua functions and ${luaTypeDefinitions.size} type definitions.`
  );
  luaFunctionInfo.forEach((description, functionName) => {
    const item = new vscode.CompletionItem(functionName, vscode.CompletionItemKind.Function);
    item.detail = 'EGOSOFT Lua Function';
    item.documentation = new vscode.MarkdownString(description);
    LuaFunctionCompletionItems.push(item);
  });
}

// Fetch and parse non-standard Lua function information
async function fetchLuaFunctionInfoFromWiki(context: vscode.ExtensionContext, forceRefresh: boolean = false) {
  luaFunctionInfo = new Map(); // Clear existing data
  const config = vscode.workspace.getConfiguration('x4CodeComplete-lua');
  const loadFromWiki = config.get('loadLuaFunctionsFromWiki') || false;
  const wikiUrl = config.get('luaFunctionWikiUrl') || '';

  if (!loadFromWiki) {
    console.log('Loading Lua functions from the Community Wiki is disabled.');
    loadLuaFunctionsInfoFromLocalFiles(rootpath);
    return;
  }

  if (!forceRefresh) {
    loadParsedDataFromGlobalState(context);
    if (luaFunctionInfo.size > 0) {
      loadLuaFunctionsInfoFromLocalFiles(rootpath);
      return; // Use cached data if available
    }
  }

  try {
    https
      .get(wikiUrl, (res) => {
        let data = '';

        // Collect data chunks
        res.on('data', (chunk) => {
          data += chunk;
        });

        // Process the complete response
        res.on('end', () => {
          const dom = new JSDOM(data);
          const document = dom.window.document;

          // Locate the table using a CSS selector
          const table = document.querySelector('#xwikicontent > table');

          if (!table) {
            console.log('Table not found in the HTML document.');
            return;
          }

          // Traverse the table rows
          const rows = table.querySelectorAll('tbody > tr');
          rows.forEach((row) => {
            const cells = row.querySelectorAll('td');

            if (cells.length >= 2) {
              const versionInfo = turndownService.turndown(cells[0].innerHTML).replace(/\*\*/g, '').trim();
              const deprecated = versionInfo.toLowerCase().includes('deprecated');
              const functionInfo = turndownService
                .turndown(cells[1].innerHTML)
                .replace(/\*\*/g, '')
                .replace(/deprecated/gi, '')
                .trim();
              const notes =
                cells.length > 2 ? turndownService.turndown(cells[2].innerHTML).replace(/\*\*/g, '').trim() : '';

              // Process only the first line of functionInfo
              const firstLine = functionInfo.split('\n')[0].trim();

              // Updated regex to handle complex return types and parameters
              const functionMatch = firstLine.match(
                /^(?:(.+?)\s*=\s*)?(\w+)\s*\((.*?)\)|^(.+?)\s+(\w+)\s*\((.*?)\)|^(.+?)$/
              );
              if (functionMatch) {
                const returnData = functionMatch[1] || functionMatch[4] || 'void';
                const functionName = functionMatch[2] || functionMatch[5] || functionMatch[7];
                const parameters =
                  functionMatch[3] || functionMatch[6]
                    ? (functionMatch[3] || functionMatch[6]).split(',').map((p) => p.trim())
                    : [];

                // Build a structured description
                let description = firstLine.replace(functionName, `\`${functionName}\``);
                if (deprecated) {
                  description = `**Deprecated** ${description}`;
                }
                const otherLines = functionInfo.split('\n').slice(1).join('\n');
                if (otherLines) {
                  description += `\n***\n${otherLines}`;
                }
                if (notes) {
                  description += `\n***\nNotes: \n\n${notes}`;
                }
                if (versionInfo) {
                  description += `\n***\nVersion Info:\n\n${versionInfo}`;
                }

                // Check for duplicates and merge descriptions
                if (luaFunctionInfo.has(functionName)) {
                  const existingDescription = luaFunctionInfo.get(functionName);
                  luaFunctionInfo.set(functionName, `${existingDescription}\n\n${description}`);
                } else {
                  luaFunctionInfo.set(functionName, description);
                }
              } else {
                console.log(`Failed to parse function info: ${firstLine}`);
              }
            }
          });

          console.log(`Fetched ${luaFunctionInfo.size} Lua functions from the website.`);
          saveParsedDataToGlobalState(context); // Save the parsed data
          loadLuaFunctionsInfoFromLocalFiles(rootpath); // Load Lua definitions (types and functions) after fetching
        });
      })
      .on('error', (err) => {
        console.error(`Failed to fetch Lua function information: ${err.message}`);
        loadLuaFunctionsInfoFromLocalFiles(rootpath);
      });
  } catch (error) {
    console.error(`Unexpected error while fetching Lua function information: ${error}`);
    loadLuaFunctionsInfoFromLocalFiles(rootpath);
  }
}

export function activate(context: vscode.ExtensionContext) {
  const config = vscode.workspace.getConfiguration('x4CodeComplete-lua');
  if (!config || !validateSettings(config)) {
    return;
  }

  rootpath = config.get('unpackedFileLocation') || '';
  extensionsFolder = config.get('extensionsFolder') || '';
  exceedinglyVerbose = config.get('exceedinglyVerbose') || false;

  // Load language files
  loadLanguageFiles(rootpath, extensionsFolder);

  // Fetch Lua function information on startup
  fetchLuaFunctionInfoFromWiki(context);

  context.subscriptions.push(
    vscode.languages.registerHoverProvider(sel, {
      provideHover: async (
        document: vscode.TextDocument,
        position: vscode.Position
      ): Promise<vscode.Hover | undefined> => {
        const tPattern = /ReadText\s*\(\s*(\d+),\s*(\d+)\s*\)/g;
        // matches:
        // ReadText(1015, 7)

        const range = document.getWordRangeAtPosition(position, tPattern);
        if (range) {
          const text = document.getText(range);
          const matches = tPattern.exec(text);
          tPattern.lastIndex = 0; // Reset regex state

          if (matches && matches.length >= 3) {
            let pageId: string | undefined;
            let textId: string | undefined;
            if (matches[1] && matches[2]) {
              // {1015,7} or {1015, 7}
              pageId = matches[1];
              textId = matches[2];
            } else if (matches[3] && matches[4]) {
              // readtext.{1015}.{7}
              pageId = matches[3];
              textId = matches[4];
            } else if (matches[5] && matches[6]) {
              // page="1015" line="7"
              pageId = matches[5];
              textId = matches[6];
            }

            if (pageId && textId) {
              if (exceedinglyVerbose) {
                console.log(`Matched pattern: ${text}, pageId: ${pageId}, textId: ${textId}`);
              }
              const languageText = findLanguageText(pageId, textId);
              if (languageText) {
                const hoverText = new vscode.MarkdownString();
                hoverText.appendMarkdown('```plaintext\n');
                hoverText.appendMarkdown(languageText);
                hoverText.appendMarkdown('\n```');
                return new vscode.Hover(hoverText, range);
              }
            }
          }
        }

        // Check for non-standard Lua functions
        const wordRange = document.getWordRangeAtPosition(position, /\b(?:C\.)?\w+\b/);
        if (wordRange) {
          const word = document.getText(wordRange);
          if (luaFunctionInfo.has(word)) {
            const description = luaFunctionInfo.get(word);
            const hoverText = new vscode.MarkdownString();
            hoverText.appendMarkdown(description);
            return new vscode.Hover(hoverText, wordRange);
          }
        }

        return undefined;
      },
    })
  );

  // Register completion provider for Lua
  context.subscriptions.push(
    vscode.languages.registerCompletionItemProvider(sel, {
      provideCompletionItems(document: vscode.TextDocument, position: vscode.Position): vscode.CompletionItem[] {
        return LuaFunctionCompletionItems;
      },
    })
  );

  // React to configuration changes
  context.subscriptions.push(
    vscode.workspace.onDidChangeConfiguration((event) => {
      if (event.affectsConfiguration('x4CodeComplete-lua')) {
        console.log('Configuration changed. Reloading settings...');
        const config = vscode.workspace.getConfiguration('x4CodeComplete-lua');

        // Reload Lua functions if the reload flag is toggled
        if (event.affectsConfiguration('x4CodeComplete-lua.reloadLuaFunctionsFromWiki')) {
          if (config.get('reloadLuaFunctionsFromWiki') === true && config.get('loadLuaFunctionsFromWiki') === true) {
            console.log('Reloading Lua functions from the Community Wiki...');
            fetchLuaFunctionInfoFromWiki(context, true);
          }
          // Reset the reload flag to false after reloading
          // Left it there, because if it made twice (after set by user and reset by extension),
          // the changes is visible in configuration. It's strange, but ...
          vscode.workspace
            .getConfiguration()
            .update('x4CodeComplete-lua.reloadLuaFunctionsFromWiki', false, vscode.ConfigurationTarget.Global);
        }
        if (event.affectsConfiguration('x4CodeComplete-lua.loadLuaFunctionsFromWiki')) {
          fetchLuaFunctionInfoFromWiki(context, config.get('loadLuaFunctionsFromWiki') === true);
        }
        if (event.affectsConfiguration('x4CodeComplete-lua.unpackedFileLocation')) {
          rootpath = config.get('unpackedFileLocation');
          fetchLuaFunctionInfoFromWiki(context);
        }
      }
    })
  );
}

// this method is called when your extension is deactivated
export function deactivate() {}
