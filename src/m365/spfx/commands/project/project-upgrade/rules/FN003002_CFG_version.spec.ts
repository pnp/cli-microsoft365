import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN003002_CFG_version } from './FN003002_CFG_version';

describe('FN003002_CFG_version', () => {
  let findings: Finding[];
  let rule: FN003002_CFG_version;

  beforeEach(() => {
    findings = [];
    rule = new FN003002_CFG_version('2.0');
  })

  it('doesn\'t return notification if schema is already up-to-date', () => {
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

  it('returns notification if schema is not up-to-date', () => {
    const project: any = {
      path: '/usr/tmp',
      configJson: {
        $schema: 'test-schema'
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
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