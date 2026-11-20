import assert from 'assert';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/Finding.js';
import { FN012021_TSC_excludeDirectories } from './FN012021_TSC_excludeDirectories.js';

describe('FN012021_TSC_excludeDirectories', () => {
  let findings: Finding[];
  let rule: FN012021_TSC_excludeDirectories;

  beforeEach(() => {
    findings = [];
    rule = new FN012021_TSC_excludeDirectories({ directories: ["**/teams", "**/temp/copilot"] });
  });

  it('doesn\'t return notification if excludeDirectories is already present', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        watchOptions: {
          excludeDirectories: ["**/teams", "**/temp/copilot"]
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if tsconfig is not available', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notification if excludeDirectories is not present', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        watchOptions: {
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });
});
