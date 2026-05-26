import assert from 'assert';
import { FN002021_DEVDEP_rushstack_eslint_config } from './FN002021_DEVDEP_rushstack_eslint_config.js';

describe('FN002021_DEVDEP_rushstack_eslint_config', () => {
  let rule: FN002021_DEVDEP_rushstack_eslint_config;

  beforeEach(() => {
    rule = new FN002021_DEVDEP_rushstack_eslint_config({ supportedRange: '4.5.2' });
  });

  it('has the correct id', () => {
    assert.strictEqual(rule.id, 'FN002021');
  });
});
