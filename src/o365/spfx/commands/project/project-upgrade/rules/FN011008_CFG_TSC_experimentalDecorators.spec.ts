import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../model';
import { FN011008_CFG_TSC_experimentalDecorators } from './FN011008_CFG_TSC_experimentalDecorators';

describe('FN011008_CFG_TSC_experimentalDecorators', () => {
  let findings: Finding[];
  let rule: FN011008_CFG_TSC_experimentalDecorators;

  beforeEach(() => {
    findings = [];
    rule = new FN011008_CFG_TSC_experimentalDecorators();
  })

  it('doesn\'t return notification if schema is already up-to-date', () => {
    const project: Project = {
      path: '/usr/tmp',
      tsConfigJson: {
        compilerOptions: {
          experimentalDecorators: true,
        },
      }
    };
    rule.visit(project, findings);
    assert.equal(findings.length, 0);
  });
});