import assert from 'assert';
import fs from 'fs';
import sinon from 'sinon';
import { sinonUtil } from '../../../../../../utils/sinonUtil.js';
import { Project, TsFile } from '../../project-model/index.js';
import { Finding } from '../../report-model/Finding.js';
import { FN025003_ESLINTRCJS_rushstack_pair_react_dom_render_unmount } from './FN025003_ESLINTRCJS_rushstack_pair_react_dom_render_unmount.js';

describe('FN025003_ESLINTRCJS_rushstack_pair_react_dom_render_unmount', () => {
  let findings: Finding[];
  let rule: FN025003_ESLINTRCJS_rushstack_pair_react_dom_render_unmount;

  afterEach(() => {
    sinonUtil.restore([
      fs.existsSync,
      fs.readFileSync
    ]);
  });

  beforeEach(() => {
    rule = new FN025003_ESLINTRCJS_rushstack_pair_react_dom_render_unmount();
    findings = [];
  });

  it('doesn\'t return notification if .eslintrc.js not found', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if .eslintrc.js is found but no nodes are present', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => ``);
    const project: Project = {
      path: '/usr/tmp',
      esLintRcJs: new TsFile('foo')
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if .eslintrc.js is found but module is not present', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => `foo`);
    const project: Project = {
      path: '/usr/tmp',
      esLintRcJs: new TsFile('foo')
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('file returned is ./.eslintrc.js when found', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    const project: Project = {
      path: '/usr/tmp',
      esLintRcJs: new TsFile('foo')
    };
    rule.visit(project, findings);
    assert.strictEqual(rule.file, './.eslintrc.js');
  });

  it(`doesn't return notification if @rushstack/pair-react-dom-render-unmount is present`, () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => `export default { parserOptions: parserOptions: { tsconfigRootDir: __dirname }, { overrides: [ { rules: { '@rushstack/pair-react-dom-render-unmount': 1 } } ] } }`);
    const project: Project = {
      path: '/usr/tmp',
      esLintRcJs: new TsFile('foo')
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notification if @rushstack/pair-react-dom-render-unmount is not present', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => `module.exports = { parserOptions: parserOptions: { tsconfigRootDir: __dirname }, { overrides: [ { rules: { } } ] } }`);
    const project: Project = {
      path: '/usr/tmp',
      esLintRcJs: new TsFile('foo')
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });
});
