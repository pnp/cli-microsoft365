import * as assert from 'assert';
import * as fs from 'fs';
import * as sinon from 'sinon';
import { Finding } from '../Finding';
import { Project, TsFile } from '../../model';
import { FN016002_TS_msgraphclient_instance } from './FN016002_TS_msgraphclient_instance';
import Utils from '../../../../../../Utils';
import { TsRule } from './TsRule';

describe('FN016002_TS_msgraphclient_instance', () => {
  let findings: Finding[];
  let rule: FN016002_TS_msgraphclient_instance;
  afterEach(() => {
    Utils.restore([
      fs.existsSync,
      fs.readFileSync,
      (TsRule as any).getParentOfType
    ]);
  });

  beforeEach(() => {
    findings = [];
    rule = new FN016002_TS_msgraphclient_instance();
  });

  it('doesn\'t return notifications if no .ts files found', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notifications if specified .ts file not found', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const project: Project = {
      path: '/usr/tmp',
      tsFiles: [
        new TsFile('foo')
      ]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notifications if couldn\'t retrieve the call expression', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => `this.serviceScope.consume;`);
    const project: Project = {
      path: '/usr/tmp',
      tsFiles: [
        new TsFile('foo')
      ]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notifications if service key is a constant', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => `this.serviceScope.consume("abc");`);
    const project: Project = {
      path: '/usr/tmp',
      tsFiles: [
        new TsFile('foo')
      ]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notifications if service key is not MSGraphClient', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => `this.serviceScope.consume(ContosoClient.serviceKey);`);
    const project: Project = {
      path: '/usr/tmp',
      tsFiles: [
        new TsFile('foo')
      ]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notifications if retrieved MSGraphClient is not assigned to a variable', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => `this.serviceScope.consume(MSGraphClient.serviceKey);`);
    const project: Project = {
      path: '/usr/tmp',
      tsFiles: [
        new TsFile('foo')
      ]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});