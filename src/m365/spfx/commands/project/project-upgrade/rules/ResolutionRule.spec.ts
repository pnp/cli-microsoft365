import { ResolutionRule } from './ResolutionRule';
import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';

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

  customCondition(project: Project): boolean {
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
  })

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
});