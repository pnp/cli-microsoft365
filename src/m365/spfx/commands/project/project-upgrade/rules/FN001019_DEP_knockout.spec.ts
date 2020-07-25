import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN001019_DEP_knockout } from './FN001019_DEP_knockout';

describe('FN001019_DEP_knockout', () => {
  let findings: Finding[];
  let rule: FN001019_DEP_knockout;

  beforeEach(() => {
    findings = [];
    rule = new FN001019_DEP_knockout('3.4.0');
  });

  it('returns notification if types definitions are missing in knockout project', () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {
          '@types/react': '15.6.5'
        }
      },
      yoRcJson: {
        "@microsoft/generator-sharepoint": {
          framework: 'knockout'
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });
});