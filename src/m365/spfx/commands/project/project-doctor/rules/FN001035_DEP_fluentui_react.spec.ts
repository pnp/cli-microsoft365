import * as assert from 'assert';
import { Project } from '../../project-model';
import { Finding } from '../../report-model/Finding';
import { FN001035_DEP_fluentui_react } from './FN001035_DEP_fluentui_react';

describe('FN001035_DEP_fluentui_react', () => {
  let findings: Finding[];
  let rule: FN001035_DEP_fluentui_react;

  beforeEach(() => {
    findings = [];
    rule = new FN001035_DEP_fluentui_react('^7.199.1');
  });

  it(`returns correct description when unsupported version of @fluentui/react found`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {
          "react": "17.0.1",
          "@fluentui/react": "7.199.0"
        }
      }
    };
    rule.visit(project, findings);
    assert(findings[0].description.includes('Install supported version '));
  });
});
