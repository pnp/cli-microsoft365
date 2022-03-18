import * as assert from 'assert';
import * as fs from 'fs';
import * as sinon from 'sinon';
import { sinonUtil } from '../../../../../../utils';
import { Project } from '../../model';
import { Finding } from '../Finding';
import { FN014005_CODE_settingsfile } from './FN014005_CODE_settingsfile';

describe('FN014005_CODE_settingsfile', () => {
  let findings: Finding[];
  let rule: FN014005_CODE_settingsfile;
  afterEach(() => {
    sinonUtil.restore(fs.existsSync);
  });

  beforeEach(() => {
    findings = [];
    rule = new FN014005_CODE_settingsfile();
  });

  it('doesn\'t return notifications if vscode settings file is present', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notifications if vscode settings file is absent', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });
});