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
    cliCompletionClinkUpdateHelp = fs.readFileSync(path.join(__dirname, '..', '..', 'docs', 'docs', 'cmd', 'cli', 'completion', 'completion-clink-update.mdx'), 'utf8');
    cliCompletionClinkUpdateHelpPlain = md.md2plain(cliCompletionClinkUpdateHelp, path.join(__dirname, '..', '..', 'docs'));
    loginHelp = fs.readFileSync(path.join(__dirname, '..', '..', 'docs', 'docs', 'cmd', 'login.mdx'), 'utf8');
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
    assert(cliCompletionClinkUpdateHelpPlain.includes('https://pnp.github.io/cli-microsoft365/user-guide/completion/'));
    assert(!cliCompletionClinkUpdateHelpPlain.includes('(https://pnp.github.io/cli-microsoft365/user-guide/completion/)'));
  });

  it('keeps only label when hyperlink URL is relative', () => {
    assert(loginHelpPlain.includes('create a custom Azure AD application'));
    assert(!loginHelpPlain.includes('(../user-guide/using-own-identity.mdx)'));
  });

  it('appends URL between brackets for hyperlinks with absolute URLs', () => {
    const src = '[CLI for Microsoft 365](https://pnp.github.io/cli-microsoft365)';
    const actual = md.md2plain(src, path.join(__dirname, '..', '..', 'docs'));
    assert.strictEqual(actual, 'CLI for Microsoft 365 (https://pnp.github.io/cli-microsoft365)');
  });

  it('converts code fences', () => {
    assert(cliCompletionClinkUpdateHelpPlain.includes('cli completion clink update > m365.lua'));
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

  it('converts content tabs with code blocks', () => {
    const tabsMd = '<Tabs>\n  <TabItem value="tab1">\n    This is tab 1 content.\n  </TabItem>\n  <TabItem value="tab2">\n    This is tab 2 content.\n  </TabItem>\n  <TabItem value="tab3">\n    This is tab 3 content.\n  </TabItem>\n</Tabs>';
    const expected = 'tab1\n    This is tab 1 content.\n  tab2\n    This is tab 2 content.\n  tab3\n    This is tab 3 content.';

    const plain = md.md2plain(tabsMd, path.join(__dirname, '..', '..', 'docs'));
    assert.strictEqual(plain, expected);
  });

  it('removes frontmatter tags', () => {
    const frontmatterMd = '---\ntitle: Demo\n---';
    const expected = '';

    const plain = md.md2plain(frontmatterMd, path.join(__dirname, '..', '..', 'docs'));
    assert.strictEqual(plain, expected);
  });

  it('removes imports', () => {
    const importsMd = 'import demo from \'../demo\';\nimport \'demo.css\';';
    const expected = '';

    const plain = md.md2plain(importsMd, path.join(__dirname, '..', '..', 'docs'));
    assert.strictEqual(plain, expected);
  });

  it('escapes underscores in an md string', () => {
    const src = 'This is _italic_';
    const actual = md.escapeMd(src);
    assert.strictEqual(actual, 'This is \\_italic\\_');
  });

  it('escapes asterisks in an md string', () => {
    const src = 'This is **bold**';
    const actual = md.escapeMd(src);
    assert.strictEqual(actual, 'This is \\*\\*bold\\*\\*');
  });

  it('escapes backticks in an md string', () => {
    const src = 'This is `code`';
    const actual = md.escapeMd(src);
    assert.strictEqual(actual, 'This is \\`code\\`');
  });

  it('escapes tilde in an md string', () => {
    const src = 'This is ~strikethrough~';
    const actual = md.escapeMd(src);
    assert.strictEqual(actual, 'This is \\~strikethrough\\~');
  });

  it('escapes pipe in an md string', () => {
    const src = 'This is | pipe';
    const actual = md.escapeMd(src);
    assert.strictEqual(actual, 'This is \\| pipe');
  });

  it('escapes new line in an md string', () => {
    const src = 'This is\nnew\nline';
    const actual = md.escapeMd(src);
    assert.strictEqual(actual, 'This is<br>new<br>line');
  });

  it(`doesn't fail escaping special md characters if the specified arg is undefined`, () => {
    const actual = md.escapeMd(undefined);
    assert.strictEqual(actual, undefined);
  });
});