import * as assert from 'assert';
import { spfx } from '../utils';

describe('utils/spfx', () => {
  it('returns false for isReactProject when .yo-rc.json and package.json not collected', () => {
    assert.strictEqual(spfx.isReactProject({
      path: '/usr/tmp'
    }), false);
  });

  it('returns false for isReactProject when .yo-rc.json not collected and package.json has no dependencies', () => {
    assert.strictEqual(spfx.isReactProject({
      path: '/usr/tmp',
      packageJson: {}
    }), false);
  });

  it('returns false for isKnockoutProject when .yo-rc.json and package.json not collected', () => {
    assert.strictEqual(spfx.isKnockoutProject({
      path: '/usr/tmp'
    }), false);
  });

  it('returns false for isKnockoutProject when .yo-rc.json not collected and package.json has no dependencies', () => {
    assert.strictEqual(spfx.isKnockoutProject({
      path: '/usr/tmp',
      packageJson: {}
    }), false);
  });
});