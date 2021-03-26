import * as assert from 'assert';
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

  visit(): void {
  }
}

describe('ManifestRule', () => {
  let rule: MockManifestRule;

  beforeEach(() => {
    rule = new MockManifestRule();
  });

  it('rule has empty file', () => {
    assert.strictEqual('', rule.file);
  });
});