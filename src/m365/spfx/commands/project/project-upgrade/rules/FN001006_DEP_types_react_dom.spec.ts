import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN001006_DEP_types_react_dom } from './FN001006_DEP_types_react_dom';

describe('FN001006_DEP_types_react_dom', () => {
  let findings: Finding[];
  let rule: FN001006_DEP_types_react_dom;

  beforeEach(() => {
    findings = [];
    rule = new FN001006_DEP_types_react_dom('15.6.6');
  })

  it('returns notification if version is not up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {
          '@types/react-dom': '15.6.5'
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });
});