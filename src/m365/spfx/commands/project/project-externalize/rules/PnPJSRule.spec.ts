import * as assert from 'assert';
import { Project } from '../../model';
import { PnPJsRule } from './PnPJsRule';
import Utils from '../../../../../../Utils';
import * as fs from 'fs';
import * as sinon from 'sinon';

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
    assert.strictEqual(findings.entries.length, 1);
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
    assert.strictEqual(findings.entries.length, 0);
  });

  it('doesnt return a shadow require when the type of component is not recognized', async () => {
    const project: Project = {
      path: 'src/m365/spfx/commands/project/test-projects/spfx-182-webpart-react',
      packageJson: {
        dependencies: {
          '@pnp/pnpjs': '1.3.5'
        }
      }
    };
    const originalExistSync = fs.existsSync;
    sinon.stub(fs, 'existsSync').callsFake((path) => {
      if (path.toString().indexOf('WebPart') > -1) {
        return false;
      }
      else {
        return originalExistSync(path);
      }
    });
    const findings = await rule.visit(project);
    assert.strictEqual(findings.suggestions.length, 0);
  });
  afterEach(() => {
    Utils.restore([
      fs.existsSync,
    ]);
  });
});