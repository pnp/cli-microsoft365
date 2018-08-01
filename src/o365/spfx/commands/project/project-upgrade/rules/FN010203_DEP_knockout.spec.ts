import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../model';
import { FN010203_DEP_knockout } from './FN010203_DEP_knockout';

describe('FN010203_DEP_knockout', () => {
  let findings: Finding[];
  let rule: FN010203_DEP_knockout;

  beforeEach(() => {
    findings = [];
    rule = new FN010203_DEP_knockout('3.4.0');
  });

  it('returns notification if types definitions are missing', () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {
          '@types/react': '15.6.5'
        }
      }
    };
    rule.visit(project, findings);
    assert.equal(findings.length, 1);
  });
});