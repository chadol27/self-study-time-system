import { readFile, readdir } from 'node:fs/promises';
import path from 'node:path';
import js from '@eslint/js';
import { Linter } from 'eslint';
import globals from 'globals';
import { HtmlValidate } from 'html-validate';

const root = path.resolve(import.meta.dirname, '..');
const sourceDirectory = path.join(root, 'src');
const htmlFiles = (await readdir(sourceDirectory)).filter((file) => file.endsWith('.html')).sort();
const sources = new Map(await Promise.all(htmlFiles.map(async (file) => [
  file,
  await readFile(path.join(sourceDirectory, file), 'utf8')
])));

function replaceScriptlets(source) {
  return source
    .replace(/<\?!=?\s*include\([^]*?\)\s*;?\s*\?>/g, '')
    .replace(/<\?=\s*appUrl\s*\?>/g, '/')
    .replace(/<\?[^]*?\?>/g, 'null');
}

function extractScripts(source) {
  return Array.from(source.matchAll(/<script(?:\s[^>]*)?>([^]*?)<\/script>/gi), (match) => match[1]);
}

function validationDocument(file, source) {
  if (file === 'Client.html') return `<!doctype html><html lang="ko"><head><title>Client</title></head><body>${source}</body></html>`;
  if (file === 'Styles.html') return `<!doctype html><html lang="ko"><head><title>Styles</title>${source}</head><body></body></html>`;
  return source;
}

const htmlValidate = new HtmlValidate({
  extends: ['html-validate:recommended'],
  rules: {
    // HtmlService sets the title after template evaluation; lowercase doctype is valid HTML.
    'doctype-style': 'off',
    'element-required-content': 'off'
  }
});
const linter = new Linter();
let failed = false;

for (const [file, source] of sources) {
  const report = await htmlValidate.validateString(
    validationDocument(file, replaceScriptlets(source)),
    path.join('src', file)
  );
  for (const result of report.results) {
    for (const message of result.messages) {
      failed = failed || message.severity === 2;
      console.error(`${result.filePath}:${message.line}:${message.column} ${message.message} (${message.ruleId})`);
    }
  }
}

const clientScript = extractScripts(replaceScriptlets(sources.get('Client.html') || '')).join('\n');
const javascriptConfig = {
  languageOptions: {
    ecmaVersion: 2022,
    sourceType: 'script',
    globals: {
      ...globals.browser,
      google: 'readonly'
    }
  },
  rules: {
    ...js.configs.recommended.rules,
    'no-unused-vars': 'off'
  }
};

for (const [file, source] of sources) {
  if (file === 'Client.html' || file === 'Styles.html') continue;
  const pageScript = extractScripts(replaceScriptlets(source)).join('\n');
  const messages = linter.verify(`${clientScript}\n${pageScript}`, javascriptConfig, `src/${file}.js`);
  for (const message of messages) {
    failed = failed || message.severity === 2;
    console.error(`src/${file}:${message.line}:${message.column} ${message.message} (${message.ruleId || 'syntax'})`);
  }
}

if (failed) process.exitCode = 1;
