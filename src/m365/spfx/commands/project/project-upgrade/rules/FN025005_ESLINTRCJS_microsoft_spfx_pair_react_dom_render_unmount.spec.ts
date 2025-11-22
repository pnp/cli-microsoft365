import assert from 'assert';
import fs from 'fs';
import sinon from 'sinon';
import { sinonUtil } from '../../../../../../utils/sinonUtil.js';
import { Project, TsFile } from '../../project-model/index.js';
import { Finding } from '../../report-model/Finding.js';
import { FN025005_ESLINTRCJS_microsoft_spfx_pair_react_dom_render_unmount } from './FN025005_ESLINTRCJS_microsoft_spfx_pair_react_dom_render_unmount.js';

describe('FN025005_ESLINTRCJS_microsoft_spfx_pair_react_dom_render_unmount', () => {
  let findings: Finding[];
  let rule: FN025005_ESLINTRCJS_microsoft_spfx_pair_react_dom_render_unmount;

  afterEach(() => {
    sinonUtil.restore([
      fs.existsSync,
      fs.readFileSync
    ]);
  });

  beforeEach(() => {
    rule = new FN025005_ESLINTRCJS_microsoft_spfx_pair_react_dom_render_unmount();
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

  it(`doesn't return notification if @microsoft/spfx/pair-react-dom-render-unmount property is absent`, () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => `export default { parserOptions: parserOptions: { tsconfigRootDir: __dirname }, { overrides: [ { rules: { } } ] } }`);
    const project: Project = {
      path: '/usr/tmp',
      esLintRcJs: new TsFile('foo')
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notification if @microsoft/spfx/pair-react-dom-render-unmount property is present', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => `module.exports = { parserOptions: parserOptions: { tsconfigRootDir: __dirname }, { overrides: [ { rules: { '@microsoft/spfx/pair-react-dom-render-unmount': 1 } } ] } }`);
    const project: Project = {
      path: '/usr/tmp',
      esLintRcJs: new TsFile('foo')
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });
});
