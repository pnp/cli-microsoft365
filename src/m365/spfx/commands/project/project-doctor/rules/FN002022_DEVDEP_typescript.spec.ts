import assert from 'assert';
import { FN002022_DEVDEP_typescript } from './FN002022_DEVDEP_typescript.js';
import { Finding } from '../../report-model/Finding.js';

describe.only('FN002022_DEVDEP_typescript', () => {
  let findings: Finding[];
  let rule: FN002022_DEVDEP_typescript;

  beforeEach(() => {
    findings = [];
    rule = new FN002022_DEVDEP_typescript('~5.3.3');
  });

  it('has the correct id', () => {
    assert.strictEqual(rule.id, 'FN002022');
  });

  it('has a description', () => {
    assert.notStrictEqual(rule.description, null);
  });

  it('does not return finding when typescript is ~5.3.3 in dependencies', () => {
    const project: any = {
      packageJson: {
        dependencies: {
          typescript: '~5.3.3'
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('does not return finding when typescript is ~5.3.3 in devDependencies', () => {
    const project: any = {
      packageJson: {
        devDependencies: {
          typescript: '~5.3.3'
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 0);
  });

  it('returns finding when typescript is lower than ~5.3.3 in dependencies', () => {
    const project: any = {
      packageJson: {
        dependencies: {
          typescript: '~5.2.0'
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
    assert.strictEqual(findings[0].occurrences[0].resolution, 'installDev typescript@~5.3.3');
  });

  it('returns finding when typescript is missing', () => {
    const project: any = {
      packageJson: {
        dependencies: {}
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
    assert.strictEqual(findings[0].occurrences[0].resolution, 'installDev typescript@~5.3.3');
  });

  it('returns finding when typescript is missing in both dependencies and devDependencies', () => {
    const project: any = {
      packageJson: {}
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
    assert.strictEqual(findings[0].occurrences[0].resolution, 'installDev typescript@~5.3.3');
  });

  it('returns finding when typescript is lower than ~5.3.3 in devDependencies', () => {
    const project: any = {
      packageJson: {
        devDependencies: {
          typescript: '~5.2.0'
        }
      }
    };
    rule.visit(project, findings);
    assert.strictEqual(findings.length, 1);
    assert.strictEqual(findings[0].occurrences[0].resolution, 'installDev typescript@~5.3.3');
  });
});