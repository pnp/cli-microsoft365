import assert from 'assert';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/index.js';
import { FN027002_OVERRIDES_qs } from './FN027002_OVERRIDES_qs.js';

describe('FN027002_OVERRIDES_qs', () => {
  let findings: Finding[];
  let rule: FN027002_OVERRIDES_qs;

  beforeEach(() => {
    findings = [];
    rule = new FN027002_OVERRIDES_qs({ version: '>=6.15.2' });
  });

  it(`doesn't return notification if package.json is not available`, () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it(`returns notification if overrides property is not defined`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {}
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it(`returns notification if overrides.qs property is not defined`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        overrides: {}
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it(`returns notification and extra remove notification if overrides.qs property is different than expected`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        overrides: {
          'qs': '0.0.1'
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 2);
    assert.strictEqual(findings[0].id, 'FN027002_REMOVE');
    assert.strictEqual(findings[0].occurrences[0].resolution, 'removeOverride overrides.qs');
    assert.strictEqual(findings[1].id, 'FN027002');
  });

  it(`returns no remove notification when overrides.qs is already at the target version`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        overrides: {
          'qs': '>=6.15.2'
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it(`returns correct node when overrides.qs is set to a string`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        overrides: {
          'qs': '0.0.1'
        },
        source: JSON.stringify({
          overrides: {
            'qs': '0.0.1'
          }
        }, null, 2)
      }
    };
    rule.visit(project, findings);
    const updateFinding = findings.find(f => f.id === 'FN027002');
    assert.strictEqual(updateFinding!.occurrences[0].position?.line, 3);
  });
});
