import assert from 'assert';
import { PackageJson, Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/index.js';
import { PackageRule } from './PackageRule.js';

class ResRule extends PackageRule {
  constructor() {
    super({ propertyName: 'main', add: true, propertyValue: '1.0.0' });
  }

  get id(): string {
    return 'FN000000';
  }
}

class ResRule2 extends PackageRule {
  constructor() {
    super({ propertyName: 'main', add: false });
  }

  get id(): string {
    return 'FN000000';
  }
}

describe('PackageRule', () => {
  let packageRule: ResRule;
  let packageRule2: ResRule;
  let findings: Finding[];

  before(() => {
    packageRule = new ResRule();
    packageRule2 = new ResRule2();
  });

  beforeEach(() => {
    findings = [];
  });

  it('returns notification if package.json does not have property', () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
      } as PackageJson
    };
    packageRule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it('doesn\'t return any notifications package.json already has the property', () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        main: "abc"
      } as any
    };
    packageRule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if the package.json property is removed already', () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {

      } as PackageJson
    };
    packageRule2.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notification if the package.json property has to be removed', () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        main: "abc",
        source: JSON.stringify({
          main: "abc"
        }, null, 2)
      } as any
    };
    packageRule2.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of findings');
    assert.strictEqual(findings[0].occurrences[0].position?.line, 2, 'Incorrect line number');
  });

  it('doesn\'t return notification if the package.json is missing', () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: undefined
    };
    packageRule2.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});
