import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN006004_CFG_PS_developer } from './FN006004_CFG_PS_developer';

describe('FN006004_CFG_PS_developer', () => {
  let findings: Finding[];
  let rule: FN006004_CFG_PS_developer;

  beforeEach(() => {
    findings = [];
    rule = new FN006004_CFG_PS_developer();
  });

  it('doesn\'t return notification if package-solution.json is not available', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});