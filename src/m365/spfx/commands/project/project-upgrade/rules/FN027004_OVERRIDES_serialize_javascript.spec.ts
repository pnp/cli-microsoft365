import assert from 'assert';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/index.js';
import { FN027004_OVERRIDES_serialize_javascript } from './FN027004_OVERRIDES_serialize_javascript.js';

describe('FN027004_OVERRIDES_serialize_javascript', () => {
  let findings: Finding[];
  let rule: FN027004_OVERRIDES_serialize_javascript;

  beforeEach(() => {
    findings = [];
    rule = new FN027004_OVERRIDES_serialize_javascript({ version: '^7.1.3' });
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

  it(`returns notification if overrides.serialize-javascript property is not defined`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        overrides: {}
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it(`returns notification and extra remove notification if overrides.serialize-javascript property is different than expected`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        overrides: {
          'serialize-javascript': '0.0.1'
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 2);
    assert.strictEqual(findings[0].id, 'FN027004_REMOVE');
    assert.strictEqual(findings[0].occurrences[0].resolution, 'removeOverride overrides.serialize-javascript');
    assert.strictEqual(findings[1].id, 'FN027004');
  });

  it(`returns no remove notification when overrides.serialize-javascript is already at the target version`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        overrides: {
          'serialize-javascript': '^7.1.3'
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it(`returns no notification when overrides.serialize-javascript is a compliant exact version`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        overrides: {
          'serialize-javascript': '7.1.4'
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it(`returns correct node when overrides.serialize-javascript is set to a string`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        overrides: {
          'serialize-javascript': '0.0.1'
        },
        source: JSON.stringify({
          overrides: {
            'serialize-javascript': '0.0.1'
          }
        }, null, 2)
      }
    };
    rule.visit(project, findings);
    const updateFinding = findings.find(f => f.id === 'FN027004');
    assert.strictEqual(updateFinding!.occurrences[0].position?.line, 3);
  });
});
