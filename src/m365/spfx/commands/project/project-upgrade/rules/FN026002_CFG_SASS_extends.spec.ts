import assert from 'assert';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/index.js';
import { FN026002_CFG_SASS_extends } from './FN026002_CFG_SASS_extends.js';

describe('FN026002_CFG_SASS_extends', () => {
  let findings: Finding[];
  let rule: FN026002_CFG_SASS_extends;

  beforeEach(() => {
    findings = [];
    rule = new FN026002_CFG_SASS_extends({ _extends: '@microsoft/spfx-web-build-rig/profiles/default/config/sass.json' });
  });

  it(`doesn't return notification if sass.json is not available`, () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it(`returns notification if extends property is not defined`, () => {
    const project: Project = {
      path: '/usr/tmp',
      sassJson: {}
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it(`returns notification if extends property is different than expected`, () => {
    const project: Project = {
      path: '/usr/tmp',
      sassJson: {
        extends: 'something'
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it(`returns correct node when extends is set to a string`, () => {
    const project: Project = {
      path: '/usr/tmp',
      sassJson: {
        extends: 'something',
        source: JSON.stringify({
          extends: 'something'
        }, null, 2)
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings[0].occurrences[0].position?.line, 2);
  });
});
