import * as assert from 'assert';
import * as fs from 'fs';
import * as sinon from 'sinon';
import { sinonUtil } from '../../../../../../utils';
import { Project, TsFile } from '../../project-model';
import { Finding } from '../../report-model/Finding';
import { FN025001_ESLINTRCJS_overrides } from './FN025001_ESLINTRCJS_overrides';

describe('FN025001_ESLINTRCJS_overrides', () => {
  let findings: Finding[];
  let rule: FN025001_ESLINTRCJS_overrides;

  afterEach(() => {
    sinonUtil.restore([
      fs.existsSync,
      fs.readFileSync
    ]);
  });

  beforeEach(() => {
    rule = new FN025001_ESLINTRCJS_overrides('{ foo: bar }');
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

  it('doesn\'t return notification if overrides property is present', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => `module.exports = { parserOptions: parserOptions: { tsconfigRootDir: __dirname }, { overrides: [ { } ] } }`);
    const project: Project = {
      path: '/usr/tmp',
      esLintRcJs: new TsFile('foo')
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('does return notification if overrides property is not present', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => `module.exports = { parserOptions: parserOptions: { tsconfigRootDir: __dirname } }`);
    const project: Project = {
      path: '/usr/tmp',
      esLintRcJs: new TsFile('foo')
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it('returns resolution for finding if overrides property is not present', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => `module.exports = { parserOptions: parserOptions: { tsconfigRootDir: __dirname } }`);
    const project: Project = {
      path: '/usr/tmp',
      esLintRcJs: new TsFile('foo')
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
    assert.strictEqual(rule.resolution, 'module.exports = {\n      overrides: [\n        { foo: bar }\n      ]\n    };');
  });

  it('does not return resolution for finding if overrides property is present', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => `module.exports = { parserOptions: parserOptions: { tsconfigRootDir: __dirname }, { overrides: [ { } ] } }`);
    const project: Project = {
      path: '/usr/tmp',
      esLintRcJs: new TsFile('foo')
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});
