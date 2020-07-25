import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN001008_DEP_react } from './FN001008_DEP_react';

describe('FN001008_DEP_react', () => {
  let findings: Finding[];
  let rule: FN001008_DEP_react;

  beforeEach(() => {
    findings = [];
    rule = new FN001008_DEP_react('15.6.2');
  })

  it('returns notification if version is not up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {
          'react': '15.6.1'
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });
});