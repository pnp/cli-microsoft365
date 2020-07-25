import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN001021_DEP_microsoft_sp_property_pane } from './FN001021_DEP_microsoft_sp_property_pane';

describe('FN001021_DEP_microsoft_sp_property_pane', () => {
  let findings: Finding[];
  let rule: FN001021_DEP_microsoft_sp_property_pane;

  beforeEach(() => {
    findings = [];
    rule = new FN001021_DEP_microsoft_sp_property_pane('15.6.6');
  })

  it('returns notification if version is not up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {
          '@microsoft/sp-property-pane': '15.6.5'
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });
});