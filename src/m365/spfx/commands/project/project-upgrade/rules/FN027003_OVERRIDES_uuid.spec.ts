import assert from 'assert';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/index.js';
import { FN027003_OVERRIDES_uuid } from './FN027003_OVERRIDES_uuid.js';

describe('FN027003_OVERRIDES_uuid', () => {
  let findings: Finding[];
  let rule: FN027003_OVERRIDES_uuid;

  beforeEach(() => {
    findings = [];
    rule = new FN027003_OVERRIDES_uuid({ version: '>=11.1.1' });
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

  it(`returns notification if overrides.uuid property is not defined`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        overrides: {}
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it(`returns notification and extra remove notification if overrides.uuid property is different than expected`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        overrides: {
          'uuid': '0.0.1'
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 2);
    assert.strictEqual(findings[0].id, 'FN027003_REMOVE');
    assert.strictEqual(findings[0].occurrences[0].resolution, 'removeOverride overrides.uuid');
    assert.strictEqual(findings[1].id, 'FN027003');
  });

  it(`returns no remove notification when overrides.uuid is already at the target version`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        overrides: {
          'uuid': '>=11.1.1'
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it(`returns correct node when overrides.uuid is set to a string`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        overrides: {
          'uuid': '0.0.1'
        },
        source: JSON.stringify({
          overrides: {
            'uuid': '0.0.1'
          }
        }, null, 2)
      }
    };
    rule.visit(project, findings);
    const updateFinding = findings.find(f => f.id === 'FN027003');
    assert.strictEqual(updateFinding!.occurrences[0].position?.line, 3);
  });
});
