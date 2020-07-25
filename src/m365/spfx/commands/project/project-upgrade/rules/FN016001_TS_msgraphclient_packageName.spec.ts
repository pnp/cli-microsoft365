import * as assert from 'assert';
import * as fs from 'fs';
import * as sinon from 'sinon';
import { Finding } from '../Finding';
import { Project, TsFile } from '../../model';
import { FN016001_TS_msgraphclient_packageName } from './FN016001_TS_msgraphclient_packageName';
import Utils from '../../../../../../Utils';
import { TsRule } from './TsRule';

describe('FN016001_TS_msgraphclient_packageName', () => {
  let findings: Finding[];
  let rule: FN016001_TS_msgraphclient_packageName;
  afterEach(() => {
    Utils.restore([
      fs.existsSync,
      fs.readFileSync,
      (TsRule as any).getParentOfType
    ]);
  });

  beforeEach(() => {
    findings = [];
    rule = new FN016001_TS_msgraphclient_packageName('@microsoft/sp-http');
  });

  it('returns empty resolution', () => {
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

  it('doesn\'t return notifications if couldn\'t retrieve the import declaration', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => `import { MSGraphClient } from '@microsoft/sp-http';`);
    sinon.stub(TsRule as any, 'getParentOfType').callsFake(() => undefined);
    const project: Project = {
      path: '/usr/tmp',
      tsFiles: [
        new TsFile('foo')
      ]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notifications if MSGraphClient is already imported from the correct package', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => `import { MSGraphClient } from '@microsoft/sp-http';`);
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