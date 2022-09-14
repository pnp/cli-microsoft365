import * as assert from 'assert';
import { Project } from '../../project-model';
import { Finding } from '../../report-model';
import { FN021003_PKG_engines_node } from './FN021003_PKG_engines_node';

describe('FN021003_PKG_engines_node', () => {
  let findings: Finding[];
  let rule: FN021003_PKG_engines_node;

  beforeEach(() => {
    findings = [];
    rule = new FN021003_PKG_engines_node('>=16.13.0 <17.0.0');
  });

  it(`doesn't return notification if package.json is not available`, () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it(`returns notification if engines property is not defined`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {}
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it(`returns notification if engines.node property is not defined`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        engines: {}
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it(`returns notification if engines.node property is different than expected`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        engines: {
          node: '16'
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });
});