import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { ManifestRule } from './ManifestRule';

class MockManifestRule extends ManifestRule {
  get id(): string {
    return 'FN000000';
  }

  get title(): string {
    return 'Mock rule';
  }

  get description(): string {
    return 'Mock manifest rule';
  }

  get resolution(): string {
    return '';
  }

  get severity(): string {
    return 'Required';
  }

  visit(project: Project, findings: Finding[]): void {
  }
}

describe('ManifestRule', () => {
  let rule: MockManifestRule;

  beforeEach(() => {
    rule = new MockManifestRule();
  })

  it('rule has empty file', () => {
    assert.strictEqual('', rule.file);
  });
});