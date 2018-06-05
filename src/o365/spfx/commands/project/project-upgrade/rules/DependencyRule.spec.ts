import { DependencyRule } from './DependencyRule';
import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../model';

class DepRule extends DependencyRule {
  constructor() {
    super('test-package', '1.0.0');
  }

  get id(): string {
    return 'FN000000';
  }
}

class DevDepRule extends DependencyRule {
  constructor() {
    super('test-package', '1.0.0', true);
  }

  get id(): string {
    return 'FN000000';
  }
}

describe('DependencyRule', () => {
  let depRule: DepRule;
  let devDepRule: DevDepRule;
  let findings: Finding[];

  before(() => {
    depRule = new DepRule();
    devDepRule = new DevDepRule();
  });

  beforeEach(() => {
    findings = [];
  })

  it('doesn\'t return any notifications if package.json not found', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    depRule.visit(project, findings);
    assert.equal(findings.length, 0);
  });

  it('returns notifications if no dev dependencies defined but dev dependency required', () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {}
      }
    };
    devDepRule.visit(project, findings);
    assert.equal(findings.length, 1);
  });

  it('doesn\'t return notification if dependency is already up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {
          'test-package': '1.0.0'
        },
        devDependencies: {}
      }
    };
    depRule.visit(project, findings);
    assert.equal(findings.length, 0);
  });
});