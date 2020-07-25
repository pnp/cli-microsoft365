import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN001005_DEP_types_react } from './FN001005_DEP_types_react';

describe('FN001005_DEP_types_react', () => {
  let findings: Finding[];
  let rule: FN001005_DEP_types_react;

  beforeEach(() => {
    findings = [];
    rule = new FN001005_DEP_types_react('15.6.6');
  })

  it('returns notification if version is not up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {
          '@types/react': '15.6.5'
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });
});