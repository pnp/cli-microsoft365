import { DependencyRule } from './DependencyRule';
import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';

class DepRule extends DependencyRule {
  constructor() {
    super('test-package', '1.0.0');
  }

  get id(): string {
    return 'FN000000';
  }
}

class DepRule2 extends DependencyRule {
  constructor() {
    super('test-package', '1.0.1');
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

class DevDepRule2 extends DependencyRule {
  constructor() {
    super('test-package', '1.0.0', true, false, false);
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
    assert.strictEqual(findings.length, 0);
  });

  it('returns notifications if no dev dependencies defined but dev dependency required', () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {}
      }
    };
    devDepRule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
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
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if newer dependency already installed (major)', () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {
          'test-package': '2.0.0'
        },
        devDependencies: {}
      }
    };
    depRule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if newer dependency already installed (minor)', () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {
          'test-package': '1.1.0'
        },
        devDependencies: {}
      }
    };
    depRule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if newer dependency already installed (patch)', () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {
          'test-package': '1.0.1'
        },
        devDependencies: {}
      }
    };
    depRule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notification even if version range satisfies package requirement', () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {
          'test-package': '>=0.0.8 <1.1.0'
        },
        devDependencies: {}
      }
    };
    depRule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it('returns notification even if semver version satisfies package requirement', () => {
    const depRule2 = new DepRule2();
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {
          'test-package': '~1.0.0'
        },
        devDependencies: {}
      }
    };
    depRule2.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it('doesn\'t return notification if the current version is invalid', () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {
          'test-package': 'github:test/test'
        },
        devDependencies: {}
      }
    };
    depRule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns uninstall resolution for uninstall a dev dependency', () => {
    const rule: DependencyRule = new DevDepRule2();
    assert.strictEqual(rule.resolution, 'uninstallDev test-package');
  });
});