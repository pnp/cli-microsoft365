import * as assert from 'assert';
import { Finding } from '../Finding';
import { Project } from '../../model';
import { ScssRule } from './ScssRule';

class MockScssRule extends ScssRule {
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

describe('ScssRule', () => {
  let rule: MockScssRule;

  beforeEach(() => {
    rule = new MockScssRule();
  })

  it('rule has empty file', () => {
    assert.strictEqual('', rule.file);
  });

  it('returns resolution type of scss', () => {
    assert.strictEqual('scss', rule.resolutionType);
  });
});