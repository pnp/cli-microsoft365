import assert from 'assert';
import fs from 'fs';
import sinon from 'sinon';
import { sinonUtil } from '../../../../../../utils/sinonUtil.js';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/Finding.js';
import { FN014005_CODE_settingsfile } from './FN014005_CODE_settingsfile.js';

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
