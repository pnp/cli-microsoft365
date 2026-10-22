import assert from 'assert';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/index.js';
import { FN027001_OVERRIDES_rushstack_heft } from './FN027001_OVERRIDES_rushstack_heft.js';

describe('FN027001_OVERRIDES_rushstack_heft', () => {
  let findings: Finding[];
  let rule: FN027001_OVERRIDES_rushstack_heft;

  beforeEach(() => {
    findings = [];
    rule = new FN027001_OVERRIDES_rushstack_heft({ version: '0.7.36' });
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

  it(`returns notification if overrides.@rushstack/heft property is not defined`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        overrides: {}
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it(`returns notification and extra remove notification if overrides.@rushstack/heft property is different than expected`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        overrides: {
          '@rushstack/heft': '0.0.1'
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 2);
    assert.strictEqual(findings[0].id, 'FN027001_REMOVE');
    assert.strictEqual(findings[0].occurrences[0].resolution, 'removeOverride overrides.@rushstack/heft');
    assert.strictEqual(findings[1].id, 'FN027001');
  });

  it(`returns no remove notification when overrides.@rushstack/heft is already at the target version`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        overrides: {
          '@rushstack/heft': '0.7.36'
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it(`returns correct node when overrides.@rushstack/heft is set to a string`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        overrides: {
          '@rushstack/heft': '0.0.1'
        },
        source: JSON.stringify({
          overrides: {
            '@rushstack/heft': '0.0.1'
          }
        }, null, 2)
      }
    };
    rule.visit(project, findings);
    const updateFinding = findings.find(f => f.id === 'FN027001');
    assert.strictEqual(updateFinding!.occurrences[0].position?.line, 3);
  });
});
