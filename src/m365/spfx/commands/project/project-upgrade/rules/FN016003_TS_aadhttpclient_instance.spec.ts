import * as assert from 'assert';
import * as fs from 'fs';
import * as sinon from 'sinon';
import { Finding } from '../Finding';
import { Project, TsFile } from '../../model';
import { FN016003_TS_aadhttpclient_instance } from './FN016003_TS_aadhttpclient_instance';
import Utils from '../../../../../../Utils';
import { TsRule } from './TsRule';

describe('FN016003_TS_aadhttpclient_instance', () => {
  let findings: Finding[];
  let rule: FN016003_TS_aadhttpclient_instance;
  afterEach(() => {
    Utils.restore([
      fs.existsSync,
      fs.readFileSync,
      (TsRule as any).getParentOfType
    ]);
  });

  beforeEach(() => {
    findings = [];
    rule = new FN016003_TS_aadhttpclient_instance();
  });

  it('returns empty resolution by default', () => {
    assert.strictEqual(rule.resolution, '');
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

  it('doesn\'t return notifications if AadHttpClient not assigned to a variable', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => `new AadHttpClient();`);
    const project: Project = {
      path: '/usr/tmp',
      tsFiles: [
        new TsFile('foo')
      ]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('uses a comment as resource when AadHttpClient created with one argument', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => `const client = new AadHttpClient(this.context.serviceScope);`);
    const project: Project = {
      path: '/usr/tmp',
      tsFiles: [
        new TsFile('foo')
      ]
    };
    rule.visit(project, findings);
    assert(findings[0].occurrences[0].resolution.indexOf('/* your resource */') > -1);
  });
});