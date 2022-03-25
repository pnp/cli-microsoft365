import * as assert from 'assert';
import { Project } from '../../project-model';
import { Finding } from '../../report-model';
import { ResolutionRule } from './ResolutionRule';

class ResRule extends ResolutionRule {
  constructor() {
    super('test-package', '1.0.0');
  }

  get id(): string {
    return 'FN000000';
  }
}

class ResRule2 extends ResolutionRule {
  constructor() {
    super('test-package', '1.0.0');
  }

  get id(): string {
    return 'FN000000';
  }

  customCondition(): boolean {
    return true;
  }
}

describe('ResolutionRule', () => {
  let depRule: ResRule;
  let depRule2: ResRule;
  let findings: Finding[];

  before(() => {
    depRule = new ResRule();
    depRule2 = new ResRule2();
  });

  beforeEach(() => {
    findings = [];
  });

  it('doesn\'t return any notifications if package.json not found', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    depRule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return any notifications if the custom condition fails', () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {}
      }
    };
    depRule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if the resolution is already up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {},
        resolutions: {
          'test-package': '1.0.0'
        }
      }
    };
    depRule2.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notification if the resolution is not up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {},
        resolutions: {
          'test-package': '0.9.0'
        },
        source: JSON.stringify({
          dependencies: {},
          resolutions: {
            'test-package': '0.9.0'
          }
        }, null, 2)
      }
    };
    depRule2.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of findings');
    assert.strictEqual(findings[0].occurrences[0].position?.line, 4, 'Incorrect line number');
  });
});