import * as assert from 'assert';
import * as fs from 'fs';
import { EOL } from 'os';
import * as path from 'path';
import { md } from './md';

describe('utils/md', () => {
  let cliCompletionClinkUpdateHelp: string;
  let cliCompletionClinkUpdateHelpPlain: string;
  let loginHelp: string;
  let loginHelpPlain: string;

  before(() => {
    cliCompletionClinkUpdateHelp = fs.readFileSync(path.join(__dirname, '..', '..', 'docs', 'docs', 'cmd', 'cli', 'completion', 'completion-clink-update.md'), 'utf8');
    cliCompletionClinkUpdateHelpPlain = md.md2plain(cliCompletionClinkUpdateHelp, path.join(__dirname, '..', '..', 'docs'));
    loginHelp = fs.readFileSync(path.join(__dirname, '..', '..', 'docs', 'docs', 'cmd', 'login.md'), 'utf8');
    loginHelpPlain = md.md2plain(loginHelp, path.join(__dirname, '..', '..', 'docs'));
  });

  it('converts title to uppercase', () => {
    assert(cliCompletionClinkUpdateHelpPlain.includes('CLI COMPLETION CLINK UPDATE'));
  });

  it('converts headings to uppercase', () => {
    assert(cliCompletionClinkUpdateHelpPlain.includes('USAGE'));
    assert(cliCompletionClinkUpdateHelpPlain.includes('OPTIONS'));
  });

  it('converts admonitions to uppercase', () => {
    assert(cliCompletionClinkUpdateHelpPlain.includes('USAGE'));
    assert(cliCompletionClinkUpdateHelpPlain.includes('OPTIONS'));
  });

  it('converts definition lists', () => {
    assert(cliCompletionClinkUpdateHelpPlain.includes('  Runs command with verbose logging'));
    assert(cliCompletionClinkUpdateHelpPlain.includes('  Runs command with debug logging'));
  });

  it('keeps only label when hyperlink label and URL are the same', () => {
    assert(cliCompletionClinkUpdateHelpPlain.includes('https://pnp.github.io/cli-microsoft365/concepts/completion/'));
    assert(!cliCompletionClinkUpdateHelpPlain.includes('(https://pnp.github.io/cli-microsoft365/concepts/completion/)'));
  });

  it('keeps only label when hyperlink URL is relative', () => {
    assert(loginHelpPlain.includes('create a custom Azure AD application'));
    assert(!loginHelpPlain.includes('(../user-guide/using-own-identity.md)'));
  });

  it('appends URL between brackets for hyperlinks with absolute URLs', () => {
    const src = '[CLI for Microsoft 365](https://pnp.github.io/cli-microsoft365)';
    const actual = md.md2plain(src, path.join(__dirname, '..', '..', 'docs'));
    assert.strictEqual(actual, 'CLI for Microsoft 365 (https://pnp.github.io/cli-microsoft365)');
  });

  it('converts code fences', () => {
    assert(cliCompletionClinkUpdateHelpPlain.includes('  cli completion clink update > m365.lua'));
    assert(!cliCompletionClinkUpdateHelpPlain.includes('```'));
  });

  it('converts inline markup', () => {
    assert(cliCompletionClinkUpdateHelpPlain.includes('m365.lua'));
    assert(!cliCompletionClinkUpdateHelpPlain.includes('`m365.lua`'));
  });

  it('removes too many empty lines', () => {
    assert(!cliCompletionClinkUpdateHelpPlain.includes(Array(5).join(EOL)));
  });

  it('includes content', () => {
    assert(cliCompletionClinkUpdateHelpPlain.includes('--verbose'));
  });
});