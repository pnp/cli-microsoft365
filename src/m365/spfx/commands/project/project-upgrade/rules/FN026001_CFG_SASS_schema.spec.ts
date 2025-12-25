import assert from 'assert';
import { Project } from '../../project-model/index.js';
import { Finding } from '../../report-model/index.js';
import { FN026001_CFG_SASS_schema } from './FN026001_CFG_SASS_schema.js';

describe('FN026001_CFG_SASS_schema', () => {
  let findings: Finding[];
  let rule: FN026001_CFG_SASS_schema;

  beforeEach(() => {
    findings = [];
    rule = new FN026001_CFG_SASS_schema({ version: 'https://developer.microsoft.com/json-schemas/heft/v0/heft-sass-plugin.schema.json' });
  });

  it(`doesn't return notification if sass.json is not available`, () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it(`returns notification if $schema property is not defined`, () => {
    const project: Project = {
      path: '/usr/tmp',
      sassJson: {}
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it(`returns notification if $schema property is different than expected`, () => {
    const project: Project = {
      path: '/usr/tmp',
      sassJson: {
        $schema: 'https://example.com/invalid-schema.json'
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it(`returns correct node when $schema is set to a string`, () => {
    const project: Project = {
      path: '/usr/tmp',
      sassJson: {
        $schema: 'https://example.com/invalid-schema.json',
        source: JSON.stringify({
          $schema: 'https://example.com/invalid-schema.json'
        }, null, 2)
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings[0].occurrences[0].position?.line, 2);
  });
});
