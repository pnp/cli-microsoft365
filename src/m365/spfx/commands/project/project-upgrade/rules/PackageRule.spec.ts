import { PackageRule } from './PackageRule';
import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project, PackageJson } from '../../model';

class ResRule extends PackageRule {
  constructor() {
    super('main', true, '1.0.0');
  }

  get id(): string {
    return 'FN000000';
  }
}

class ResRule2 extends PackageRule {
  constructor() {
    super('main', false);
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
  })

  it('returns notification if package.json does not have property', () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {

      } as PackageJson
    };
    packageRule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it('doesn\'t return any notifications package.json has the property', () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        main: "abc"
      } as any
    };
    packageRule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if the packege.json property is being removed already', () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {

      } as PackageJson
    };
    packageRule2.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notification if the packege.json property has to be removed', () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        main: "abc"
      } as any
    };
    packageRule2.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it('doesn\'t return notification if the packege.json is missing', () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: undefined
    };
    packageRule2.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});