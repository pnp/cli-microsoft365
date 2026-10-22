import assert from 'assert';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/Finding.js';
import { FN001019_DEP_knockout } from './FN001019_DEP_knockout.js';

describe('FN001019_DEP_knockout', () => {
  let findings: Finding[];
  let rule: FN001019_DEP_knockout;

  beforeEach(() => {
    findings = [];
    rule = new FN001019_DEP_knockout({ packageVersion: '3.4.0' });
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
