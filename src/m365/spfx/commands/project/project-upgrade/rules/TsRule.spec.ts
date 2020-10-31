import * as assert from 'assert';
import { TsRule } from './TsRule';

class MockTsRule extends TsRule {
  get id(): string {
    return 'FN000000';
  }

  get title(): string {
    return '';
  }

  get description(): string {
    return '';
  }

  get severity(): string {
    return 'Required';
  }

  visit(): void {

  }
}

describe('TsRule', () => {
  let rule: MockTsRule;

  beforeEach(() => {
    rule = new MockTsRule();
  })

  it('returns no resolution by default', () => {
    assert.strictEqual(rule.resolution, '');
  });

  it('returns ts resolutionType', () => {
    assert.strictEqual(rule.resolutionType, 'ts');
  });

  it('returns no file name by default', () => {
    assert.strictEqual(rule.file, '');
  });
});