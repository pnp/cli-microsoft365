import * as assert from 'assert';
import { Project } from '../../project-model';
import { Finding } from '../../report-model/Finding';
import { DependencyRule } from './DependencyRule';

class MockDepRule extends DependencyRule {
  constructor() {
    super('package', '1.0.0', false);
  }

  get id(): string {
    return 'FN000001';
  }
}

class MockDevDepRule extends DependencyRule {
  constructor() {
    super('package', '1.0.0', true);
  }

  get id(): string {
    return 'FN000002';
  }
}

describe('DependencyRule', () => {
  let findings: Finding[];
  let depRule: MockDepRule;
  let devDepRule: MockDevDepRule;

  beforeEach(() => {
    findings = [];
    depRule = new MockDepRule();
    devDepRule = new MockDevDepRule();
  });

  it('returns empty description by default', () => {
    assert.strictEqual(depRule.description, '');
  });

  it('returns install resolution for a dependency', () => {
    assert(!depRule.resolution.includes('installDev'));
    assert(depRule.resolution.includes('install'));
  });

  it('returns installDev resolution for a devDependency', () => {
    assert(devDepRule.resolution.includes('installDev'));
  });

  it(`depRule doesn't return notifications when project has no dependencies`, () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    depRule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it(`devDepRule doesn't return notifications when project has no devDependencies`, () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    devDepRule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it(`returns a missing package resolution when couldn't resolve dependency version`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {
          'package': 'invalid'
        }
      }
    };
    depRule.visit(project, findings);
    assert(findings[0].description.includes('Install missing package'));
  });

  it(`returns a missing package resolution when couldn't resolve devDependency version`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        devDependencies: {
          'package': 'invalid'
        }
      }
    };
    devDepRule.visit(project, findings);
    assert(findings[0].description.includes('Install missing package'));
  });

  it(`returns notification when installed dependency version doesn't satisfy the supported range`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {
          'package': '0.9.1'
        }
      }
    };
    depRule.visit(project, findings);
    assert(findings[0].description.includes('Install supported version'));
  });

  it(`returns notification when installed devDependency version doesn't satisfy the supported range`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        devDependencies: {
          'package': '0.9.1'
        }
      }
    };
    devDepRule.visit(project, findings);
    assert(findings[0].description.includes('Install supported version'));
  });
});