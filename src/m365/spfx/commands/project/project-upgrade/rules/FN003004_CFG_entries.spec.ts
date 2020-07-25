import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { FN003004_CFG_entries } from './FN003004_CFG_entries';

describe('FN003004_CFG_entries', () => {
  let findings: Finding[];
  let rule: FN003004_CFG_entries;

  beforeEach(() => {
    findings = [];
    rule = new FN003004_CFG_entries();
  });

  it('has empty resolution', () => {
    assert.strictEqual(rule.resolution, '');
  });

  it('returns notification if entries is present', () => {
    const project: any = {
      path: '/usr/tmp',
      configJson: {
        "entries": [
          {
            "entry": "./lib/extensions/helloWorld/HelloWorldApplicationCustomizer.js",
            "manifest": "./src/extensions/helloWorld/HelloWorldApplicationCustomizer.manifest.json",
            "outputPath": "./dist/hello-world.bundle.js"
          }]
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
  });

  it('should show the entries schema in the resolution', () => {
    const project: any = {
      path: '/usr/tmp',
      configJson: {
        "entries": [
          {
            "entry": "./lib/extensions/helloWorld/HelloWorldApplicationCustomizer.js",
            "manifest": "./src/extensions/helloWorld/HelloWorldApplicationCustomizer.manifest.json",
            "outputPath": "./dist/hello-world.bundle.js"
          }]
      }
    };
    rule.visit(project, findings);

    const resolution: any = JSON.parse(findings[0].occurrences[0].resolution);
    const entries: any = resolution.entries;

    assert.notStrictEqual(entries, undefined);
    assert.strictEqual(entries[0].entry, './lib/extensions/helloWorld/HelloWorldApplicationCustomizer.js');
    assert.strictEqual(entries[0].manifest, './src/extensions/helloWorld/HelloWorldApplicationCustomizer.manifest.json');
    assert.strictEqual(entries[0].outputPath, './dist/hello-world.bundle.js');
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
      configJson: {}
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });
});