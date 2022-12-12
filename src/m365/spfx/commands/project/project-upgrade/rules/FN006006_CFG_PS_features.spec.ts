import * as assert from 'assert';
import { Project } from '../../project-model';
import { Finding } from '../../report-model/Finding';
import { FN006006_CFG_PS_features } from './FN006006_CFG_PS_features';

describe('FN006006_CFG_PS_features', () => {
  let findings: Finding[];
  let rule: FN006006_CFG_PS_features;

  beforeEach(() => {
    findings = [];
    rule = new FN006006_CFG_PS_features();
  });

  it('returns empty resolution by default', () => {
    assert.strictEqual(rule.resolution, '');
  });

  it(`doesn't return notification if package-solution.json is not available`, () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it(`doesn't return notification if package-solution.json doesn't have solution`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageSolutionJson: {
        $schema: 'test-schema'
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it(`doesn't return notification if package-solution.json has features`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageSolutionJson: {
        $schema: 'test-schema',
        solution: {
          features: [{}]
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it(`doesn't return notification if package-solution.json doesn't have components`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageSolutionJson: {
        $schema: 'test-schema',
        solution: {}
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it(`doesn't fail when didn't retrieve package.json`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageSolutionJson: {
        $schema: '',
        solution: {}
      },
      manifests: [{
        $schema: '',
        path: '/usr/tmp/component',
        componentType: 'WebPart',
        alias: 'webpart'
      }]
    };
    rule.visit(project, findings);
    assert(findings[0].occurrences[0].resolution.indexOf(' webpart Feature') > -1);
  });

  it(`returns a feature for every component`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageSolutionJson: {
        $schema: '',
        solution: {}
      },
      manifests: [{
        $schema: '',
        path: '/usr/tmp/component',
        componentType: 'WebPart',
        alias: 'webpart'
      }, {
        $schema: '',
        path: '/usr/tmp/component2',
        componentType: 'Extension',
        alias: 'extension'
      }]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings[0].occurrences.length, 2);
  });

  it(`returns features with the ID of the matching component`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageSolutionJson: {
        $schema: '',
        solution: {}
      },
      manifests: [{
        $schema: '',
        path: '/usr/tmp/component',
        componentType: 'WebPart',
        alias: 'webpart',
        id: '41791d1c-c475-4217-b248-ed919db66bc2'
      }]
    };
    rule.visit(project, findings);
    const resolution = JSON.parse(findings[0].occurrences[0].resolution);
    assert.strictEqual(resolution.solution.features[0].id, '41791d1c-c475-4217-b248-ed919db66bc2');
  });

  it(`returns features that contain only the component they're linked to`, () => {
    const project: Project = {
      path: '/usr/tmp',
      packageSolutionJson: {
        $schema: '',
        solution: {}
      },
      manifests: [{
        $schema: '',
        path: '/usr/tmp/component',
        componentType: 'WebPart',
        alias: 'webpart',
        id: '41791d1c-c475-4217-b248-ed919db66bc2'
      }, {
        $schema: '',
        path: '/usr/tmp/component2',
        componentType: 'Extension',
        alias: 'extension',
        id: '41791d1c-c475-4217-b248-ed919db66bc3'
      }]
    };
    rule.visit(project, findings);
    const resolution = JSON.parse(findings[0].occurrences[0].resolution);
    assert.deepStrictEqual(resolution.solution.features[0].componentIds, ['41791d1c-c475-4217-b248-ed919db66bc2']);
  });
});