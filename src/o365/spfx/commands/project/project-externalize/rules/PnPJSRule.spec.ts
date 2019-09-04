import * as assert from 'assert';
import { Project } from '../../project-upgrade/model';
import { PnPJsRule } from './PnPJsRule';

describe('PnPJsRule', () => {
  let rule: PnPJsRule;

  beforeEach(() => {
    rule = new PnPJsRule();
  })

  it('returns notification if dependency is here', async () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {
          '@pnp/pnpjs': '1.3.5'
        }
      }
    };
    const findings = await rule.visit(project);
    assert.equal(findings.length, 1);
  });

  it('returns no notification if dependency is not here', async () => {
    const project: Project = {
      path: '/usr/tmp',
      packageJson: {
        dependencies: {
          '@pnp/pnpts': '1.3.5'
        }
      }
    };
    const findings = await rule.visit(project);
    assert.equal(findings.length, 0);
  });
});