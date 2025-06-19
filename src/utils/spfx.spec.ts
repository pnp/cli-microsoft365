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
    assert.strictEqual(version, '16.0.0');
  });

  it('returns correct Node version for a single version', () => {
    const version = spfx.getHighestNodeVersion('^10');
    assert.strictEqual(version, '10.x');
  });

  it('returns correct Node version for a range with multiple versions', () => {
    const version = spfx.getHighestNodeVersion('^12.13 || ^14.15 || ^16.13');
    assert.strictEqual(version, '16.13.x');
  });

  it('returns correct Node version when only minor version differ', () => {
    const version = spfx.getHighestNodeVersion('8.1 || 8.2');
    assert.strictEqual(version, '8.2.x');
  });

  it('returns highest major for disjoint ranges', () => {
    const version = spfx.getHighestNodeVersion('>=16.13.0 <17.0.0 || >=18.17.1 <19.0.0');
    assert.strictEqual(version, '18.17.1');
  });

  it('throws when range is empty', () => {
    assert.throws(() => spfx.getHighestNodeVersion(''), new Error('Node version range was not provided.'));
  });

  it('throws when range cannot be normalized', () => {
    assert.throws(() => spfx.getHighestNodeVersion('invalid-range'), new Error("Unable to resolve the highest Node version for range 'invalid-range'."));
  });

  it('throws when min version cannot be determined', () => {
    assert.throws(() => spfx.getHighestNodeVersion('invalid || >=1.0.0 <1.0.0'), new Error("Unable to resolve the highest Node version for range 'invalid || >=1.0.0 <1.0.0'."));
  });
});
