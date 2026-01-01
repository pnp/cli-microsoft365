import assert from 'assert';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/index.js';
import { FN021007_PKG_scripts_start } from './FN021007_PKG_scripts_start.js';

describe('FN021007_PKG_scripts_start', () => {
  let findings: Finding[];
  let rule: FN021007_PKG_scripts_start;
  beforeEach(() => {
    findings = [];
    rule = new FN021007_PKG_scripts_start({ script: 'heft start --clean' });
  });

  it(`doesn't return notification if package.json is not available`, () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it(`returns notification if scripts property is not defined`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {}
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it(`returns notification if scripts.start property is not defined`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        scripts: {}
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it(`returns notification if scripts.start property is different than expected`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        scripts: {
          start: 'start'
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it(`returns correct node when start is set to a string`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        scripts: {
          start: 'start'
        },
        source: JSON.stringify({
          scripts: {
            start: 'start'
          }
        }, null, 2)
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings[0].occurrences[0].position?.line, 3);
  });
});
