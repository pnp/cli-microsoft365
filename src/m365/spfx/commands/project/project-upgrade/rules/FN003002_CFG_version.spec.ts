import * as assert from 'assert';
import { Project } from '../../model';
import { Finding } from '../Finding';
import { FN003002_CFG_version } from './FN003002_CFG_version';

describe('FN003002_CFG_version', () => {
  let findings: Finding[];
  let rule: FN003002_CFG_version;

  beforeEach(() => {
    findings = [];
    rule = new FN003002_CFG_version('2.0');
  })

  it('doesn\'t return notification if version is already up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      configJson: {
        $schema: 'test-schema',
        version: '2.0',
        bundles: {}
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notification if version is not up-to-date', () => {
    const project: any = {
      path: '/usr/tmp',
      configJson: {
        $schema: 'test-schema',
        version: '1.0',
        source: JSON.stringify({
          $schema: 'test-schema',
          version: '1.0'
        }, null, 2)
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of findings');
    assert.strictEqual(findings[0].occurrences[0].position?.line, 3, 'Incorrect line number');
  });

  it('exits if no config json', () => {
    const project: Project = {
      path: '/usr/tmp',
      configJson: undefined
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});