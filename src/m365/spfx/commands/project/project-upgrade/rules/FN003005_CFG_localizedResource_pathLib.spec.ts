import * as assert from 'assert';
import { Project } from '../../project-model';
import { Finding } from '../../report-model/Finding';
import { FN003005_CFG_localizedResource_pathLib } from './FN003005_CFG_localizedResource_pathLib';

describe('FN003005_CFG_localizedResource_pathLib', () => {
  let findings: Finding[];
  let rule: FN003005_CFG_localizedResource_pathLib;

  beforeEach(() => {
    findings = [];
    rule = new FN003005_CFG_localizedResource_pathLib();
  });

  it('has empty resolution', () => {
    assert.strictEqual(rule.resolution, '');
  });

  it('doesn\'t return notification if no config.json', () => {
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if no localized resources', () => {
    const project: Project = {
      path: '/usr/tmp',
      configJson: {}
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notification if localized resource path starts with lib/', () => {
    const project: Project = {
      path: '/usr/tmp',
      configJson: {
        localizedResources: {
          'HelloWorldWebPartStrings': 'lib/webparts/helloWorld/loc/{locale}.js'
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notification if localized resource path doesn\'t start with lib/', () => {
    const project: Project = {
      path: '/usr/tmp',
      configJson: {
        localizedResources: {
          'HelloWorldWebPartStrings': 'webparts/helloWorld/loc/{locale}.js'
        },
        source: JSON.stringify({
          localizedResources: {
            'HelloWorldWebPartStrings': 'webparts/helloWorld/loc/{locale}.js'
          }
        }, null, 2)
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of findings');
    assert.strictEqual(findings[0].occurrences[0].position?.line, 3, 'Incorrect line number');
  });

  it('returned notification has correct resolution', () => {
    const project: Project = {
      path: '/usr/tmp',
      configJson: {
        localizedResources: {
          'HelloWorldWebPartStrings': 'webparts/helloWorld/loc/{locale}.js'
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings[0].occurrences[0].resolution, JSON.stringify({
      localizedResources: {
        'HelloWorldWebPartStrings': 'lib/webparts/helloWorld/loc/{locale}.js'
      }
    }, null, 2));
  });
});