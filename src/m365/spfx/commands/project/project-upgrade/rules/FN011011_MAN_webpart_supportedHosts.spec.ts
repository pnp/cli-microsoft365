import * as assert from 'assert';
import { Project } from '../../project-model';
import { Finding } from '../../report-model/Finding';
import { FN011011_MAN_webpart_supportedHosts } from './FN011011_MAN_webpart_supportedHosts';

describe('FN011011_MAN_webpart_supportedHosts', () => {
  let findings: Finding[];
  let rule: FN011011_MAN_webpart_supportedHosts;

  beforeEach(() => {
    findings = [];
  });

  it('doesn\'t return notifications if no manifests collected', () => {
    rule = new FN011011_MAN_webpart_supportedHosts(true);
    const project: Project = {
      path: '/usr/tmp'
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notifications if no supportedHosts found while it should be removed', () => {
    rule = new FN011011_MAN_webpart_supportedHosts(false);
    const project: Project = {
      path: '/usr/tmp',
      manifests: [
        {
          path: '/usr/tmp/manifest.json',
          $schema: 'test-schema',
          componentType: 'WebPart'
        }
      ]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('doesn\'t return notifications if supportedHosts found while it should be added', () => {
    rule = new FN011011_MAN_webpart_supportedHosts(true);
    const project: Project = {
      path: '/usr/tmp',
      manifests: [
        {
          path: '/usr/tmp/manifest.json',
          $schema: 'test-schema',
          componentType: 'WebPart',
          supportedHosts: ['SharePointWebPart']
        }
      ]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns notifications if supportedHosts not found while it should be added', () => {
    rule = new FN011011_MAN_webpart_supportedHosts(true);
    const project: Project = {
      path: '/usr/tmp',
      manifests: [
        {
          path: '/usr/tmp/manifest.json',
          $schema: 'test-schema',
          componentType: 'WebPart',
          source: JSON.stringify({
            path: '/usr/tmp/manifest.json',
            $schema: 'test-schema',
            componentType: 'WebPart'
          }, null, 2)
        }
      ]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of findings');
    assert.strictEqual(findings[0].occurrences[0].position?.line, 1, 'Incorrect line number');
  });

  it('returns notifications if supportedHosts found while it should be removed', () => {
    rule = new FN011011_MAN_webpart_supportedHosts(false);
    const project: Project = {
      path: '/usr/tmp',
      manifests: [
        {
          path: '/usr/tmp/manifest.json',
          $schema: 'test-schema',
          componentType: 'WebPart',
          supportedHosts: ['SharePointWebPart'],
          source: JSON.stringify({
            path: '/usr/tmp/manifest.json',
            $schema: 'test-schema',
            componentType: 'WebPart',
            supportedHosts: ['SharePointWebPart']
          }, null, 2)
        }
      ]
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1, 'Incorrect number of findings');
    assert.strictEqual(findings[0].occurrences[0].position?.line, 5, 'Incorrect line number');
  });
});