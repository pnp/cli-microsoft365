import assert from 'assert';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/Finding.js';
import { FN001020_DEP_types_knockout } from './FN001020_DEP_types_knockout.js';

describe('FN001020_DEP_types_knockout', () => {
  let findings: Finding[];
  let rule: FN001020_DEP_types_knockout;

  beforeEach(() => {
    findings = [];
    rule = new FN001020_DEP_types_knockout('3.4.39');
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
