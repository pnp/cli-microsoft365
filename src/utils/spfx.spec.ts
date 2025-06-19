import assert from 'assert';
import { spfx } from '../utils/spfx.js';

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

  it('returns correct Node version for a given range', () => {
    const version = spfx.getHighestNodeVersion('>=14.0.0 <15.0.0 || >=16.0.0 <17.0.0');
    assert.strictEqual(version, '17.0.x');
  });

  it('returns correct Node version for a single version', () => {
    const version = spfx.getHighestNodeVersion('^10');
    assert.strictEqual(version, '10.0.x');
  });

  it('returns correct Node version for a range with multiple versions', () => {
    const version = spfx.getHighestNodeVersion('^12.13 || ^14.15 || ^16.13');
    assert.strictEqual(version, '16.13.x');
  });

  it('returns correct Node version when only minor version differ', () => {
    const version = spfx.getHighestNodeVersion('8.1 || 8.2');
    assert.strictEqual(version, '8.2.x');
  });
});