import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN003003_CFG_bundles } from './FN003003_CFG_bundles';

describe('FN003003_CFG_bundles', () => {
  let findings: Finding[];
  let rule: FN003003_CFG_bundles;

  beforeEach(() => {
    findings = [];
    rule = new FN003003_CFG_bundles();
  });

  it('has empty resolution', () => {
    assert.strictEqual(rule.resolution, '');
  });

  it('doesn\'t return notification if bundles is already up-to-date', () => {
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

  it('returns notification if entries has to be converted to bundles', () => {
    const project: any = {
      path: '/usr/tmp',
      configJson: {
        "entries": [
          {
            "entry": "./lib/extensions/helloWorld/HelloWorldApplicationCustomizer.js",
            "manifest": "./src/extensions/helloWorld/HelloWorldApplicationCustomizer.manifest.json",
            "outputPath": "./dist/hello-world.bundle.js"
          }
        ]
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it('should correctly convert entries schema to bundle schema', () => {
    const project: any = {
      path: '/usr/tmp',
      configJson: {
        entries: [
          {
            entry: "./lib/extensions/helloWorld/HelloWorldApplicationCustomizer.js",
            manifest: "./src/extensions/helloWorld/HelloWorldApplicationCustomizer.manifest.json",
            outputPath: "./dist/hello-world.bundle.js"
          }
        ]
      }
    };
    rule.visit(project, findings);

    const resolution: any = JSON.parse(findings[0].occurrences[0].resolution);
    const bundle1: any = resolution.bundles["hello-world-application-customizer"];

    assert.notStrictEqual(bundle1, undefined, 'Bundle undefined');
    assert.strictEqual(bundle1.components[0].entrypoint, './lib/extensions/helloWorld/HelloWorldApplicationCustomizer.js', 'Invalid entrypoint');
    assert.strictEqual(bundle1.components[0].manifest, './src/extensions/helloWorld/HelloWorldApplicationCustomizer.manifest.json', 'Invalid manifest');
    assert.strictEqual(bundle1.components[0].outputPath, undefined, 'outputPath defined');
  });

  it('doesn\'t return notification if no config.json', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if no entries', () => {
    const project: Project = {
      path: '/usr/tmp',
      configJson: {
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});