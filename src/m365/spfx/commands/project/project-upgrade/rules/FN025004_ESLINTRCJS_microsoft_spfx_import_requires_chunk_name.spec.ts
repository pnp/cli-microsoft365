import assert from 'assert';
import fs from 'fs';
import sinon from 'sinon';
import { sinonUtil } from '../../../../../../utils/sinonUtil.js';
import { Project, TsFile } from '../../project-model/index.js';
import { Finding } from '../../report-model/Finding.js';
import { FN025004_ESLINTRCJS_microsoft_spfx_import_requires_chunk_name } from './FN025004_ESLINTRCJS_microsoft_spfx_import_requires_chunk_name.js';

describe('FN025004_ESLINTRCJS_microsoft_spfx_import_requires_chunk_name', () => {
  let findings: Finding[];
  let rule: FN025004_ESLINTRCJS_microsoft_spfx_import_requires_chunk_name;

  afterEach(() => {
    sinonUtil.restore([
      fs.existsSync,
      fs.readFileSync
    ]);
  });

  beforeEach(() => {
    rule = new FN025004_ESLINTRCJS_microsoft_spfx_import_requires_chunk_name();
    findings = [];
  });

  it('file returned is ./.eslintrc.js', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    assert.strictEqual(rule.file, './.eslintrc.js');
  });

  it(`doesn't return notification if .eslintrc.js not found`, () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it(`doesn't return notification if .eslintrc.js is found but no nodes are present`, () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => ``);
    const project: Project = {
      path: '/usr/tmp',
      esLintRcJs: new TsFile('foo')
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it(`doesn't return notification if @microsoft/spfx/import-requires-chunk-name property is absent`, () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => `export default { parserOptions: parserOptions: { tsconfigRootDir: __dirname }, { overrides: [ { rules: { } } ] } }`);
    const project: Project = {
      path: '/usr/tmp',
      esLintRcJs: new TsFile('foo')
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notification if @microsoft/spfx/import-requires-chunk-name property is present', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => `module.exports = { parserOptions: parserOptions: { tsconfigRootDir: __dirname }, { overrides: [ { rules: { '@microsoft/spfx/import-requires-chunk-name': 1 } } ] } }`);
    const project: Project = {
      path: '/usr/tmp',
      esLintRcJs: new TsFile('foo')
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });
});
