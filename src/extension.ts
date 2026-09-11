import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';
import { convertMarkdownToDocx } from './converter';
import { ConverterOptions } from './styles';

export function activate(context: vscode.ExtensionContext) {
  console.log('Markdown to DOCX Converter extension is active.');

  // Command 1: Convert Active Editor Markdown Document
  const convertActiveCmd = vscode.commands.registerCommand('markdownToDocx.convert', async () => {
    const editor = vscode.window.activeTextEditor;
    if (!editor) {
      vscode.window.showErrorMessage('No active Markdown editor found.');
      return;
    }

    const document = editor.document;
    if (document.languageId !== 'markdown' && !document.fileName.endsWith('.md')) {
      vscode.window.showErrorMessage('Active document is not a Markdown file.');
      return;
    }

    await convertFile(document.uri);
  });

  // Command 2: Convert Explorer Context Menu File
  const convertFileCmd = vscode.commands.registerCommand(
    'markdownToDocx.convertFile',
    async (fileUri?: vscode.Uri) => {
      const uri = fileUri || vscode.window.activeTextEditor?.document.uri;
      if (!uri) {
        vscode.window.showErrorMessage('No Markdown file selected for conversion.');
        return;
      }
      await convertFile(uri);
    }
  );

  context.subscriptions.push(convertActiveCmd, convertFileCmd);
}

async function convertFile(fileUri: vscode.Uri) {
  const filePath = fileUri.fsPath;
  const parsedPath = path.parse(filePath);
  const docxOutputPath = path.join(parsedPath.dir, `${parsedPath.name}.docx`);

  const config = vscode.workspace.getConfiguration('markdownToDocx');
  const options: ConverterOptions = {
    headerImage: config.get<string>('headerImage'),
    footerImage: config.get<string>('footerImage'),
    watermarkText: config.get<string>('watermarkText'),
    accentColor: config.get<string>('accentColor'),
    openAfterConversion: config.get<boolean>('openAfterConversion'),
  };

  // Resolve relative paths for images relative to workspace root or target file directory
  if (options.headerImage && !path.isAbsolute(options.headerImage)) {
    options.headerImage = path.resolve(parsedPath.dir, options.headerImage);
  }
  if (options.footerImage && !path.isAbsolute(options.footerImage)) {
    options.footerImage = path.resolve(parsedPath.dir, options.footerImage);
  }

  await vscode.window.withProgress(
    {
      location: vscode.ProgressLocation.Notification,
      title: `Converting ${parsedPath.base} to DOCX...`,
      cancellable: false,
    },
    async () => {
      try {
        const mdContent = fs.readFileSync(filePath, 'utf-8');
        const docxBuffer = await convertMarkdownToDocx(mdContent, options);

        fs.writeFileSync(docxOutputPath, docxBuffer);

        vscode.window
          .showInformationMessage(
            `Successfully converted to ${parsedPath.name}.docx!`,
            'Open DOCX',
            'Reveal in Explorer'
          )
          .then((selection) => {
            if (selection === 'Open DOCX') {
              vscode.env.openExternal(vscode.Uri.file(docxOutputPath));
            } else if (selection === 'Reveal in Explorer') {
              vscode.commands.executeCommand('revealFileInOS', vscode.Uri.file(docxOutputPath));
            }
          });
      } catch (err: any) {
        vscode.window.showErrorMessage(`Failed to convert Markdown to DOCX: ${err.message || err}`);
      }
    }
  );
}

export function deactivate() {}
