import assert from 'assert';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/Finding.js';
import { DependencyRule } from './DependencyRule.js';

class DepRule extends DependencyRule {
  constructor() {
    super({
      packageName: 'test-package',
      packageVersion: '1.0.0'
    });
  }

  get id(): string {
    return 'FN000000';
  }
}

class DepRule2 extends DependencyRule {
  constructor() {
    super({
      packageName: 'test-package',
      packageVersion: '1.0.1'
    });
  }

  get id(): string {
    return 'FN000000';
  }
}

class DevDepRule extends DependencyRule {
  constructor() {
    super({
      packageName: 'test-package',
      packageVersion: '1.0.0',
      isDevDep: true
    });
  }

  get id(): string {
    return 'FN000000';
  }
}

class DevDepRule2 extends DependencyRule {
  constructor() {
    super({
      packageName: 'test-package',
      packageVersion: '1.0.0',
      isDevDep: true,
      isOptional: false,
      add: false
    });
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
  });

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
        dependencies: {},
        source: JSON.stringify({
          dependencies: {}
        }, null, 2)
      }
    };
    devDepRule.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of findings');
    assert.strictEqual(findings[0].occurrences[0].position?.line, 1, 'Incorrect line number');
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
        devDependencies: {},
        source: JSON.stringify({
          dependencies: {
            'test-package': '>=0.0.8 <1.1.0'
          },
          devDependencies: {}
        }, null, 2)
      }
    };
    depRule.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of finding');
    assert.strictEqual(findings[0].occurrences[0].position?.line, 3, 'Incorrect line number');
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

  it('returns notification when rule version is higher than open-ended range', () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {
          'test-package': '>=0.5.0 <0.9.0'
        },
        devDependencies: {},
        source: JSON.stringify({
          dependencies: {
            'test-package': '>=0.5.0 <0.9.0'
          },
          devDependencies: {}
        }, null, 2)
      }
    };
    depRule.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of findings');
    assert.strictEqual(findings[0].occurrences[0].position?.line, 3, 'Incorrect line number');
  });

  it('handles version range with multiple upper bounds correctly', () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {
          'test-package': '>=0.2.0 <=0.3.0 || >=0.4.0 <=0.9.0'
        },
        devDependencies: {},
        source: JSON.stringify({
          dependencies: {
            'test-package': '>=0.2.0 <=0.3.0 || >=0.4.0 <=0.9.0'
          },
          devDependencies: {}
        }, null, 2)
      }
    };
    depRule.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of findings');
  });

  it('handles version range without upper bound correctly', () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {
          'test-package': '>=1.5.0'
        },
        devDependencies: {},
        source: JSON.stringify({
          dependencies: {
            'test-package': '>=1.5.0'
          },
          devDependencies: {}
        }, null, 2)
      }
    };
    depRule.visit(project, findings);
    assert.strictEqual(findings.length, 0, 'Incorrect number of findings');
  });

  it('handles version range with only upper bounds correctly', () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {
          'test-package': '<=0.9.0'
        },
        devDependencies: {},
        source: JSON.stringify({
          dependencies: {
            'test-package': '<=0.9.0'
          },
          devDependencies: {}
        }, null, 2)
      }
    };
    depRule.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of findings');
  });

  it('returns uninstall resolution for uninstall a dev dependency', () => {
    const rule: DependencyRule = new DevDepRule2();
    assert.strictEqual(rule.resolution, 'uninstallDev test-package');
  });
});
