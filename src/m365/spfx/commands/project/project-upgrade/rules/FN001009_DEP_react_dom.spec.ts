import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN001009_DEP_react_dom } from './FN001009_DEP_react_dom';

describe('FN001009_DEP_react_dom', () => {
  let findings: Finding[];
  let rule: FN001009_DEP_react_dom;

  beforeEach(() => {
    findings = [];
    rule = new FN001009_DEP_react_dom('15.6.2');
  })

  it('returns notification if version is not up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {
          'react-dom': '15.6.1'
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });
});