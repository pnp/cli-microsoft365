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
    assert.strictEqual(version, '16.x');
  });

  it('returns correct Node version for a single version', () => {
    const version = spfx.getHighestNodeVersion('^10');
    assert.strictEqual(version, '10.x');
  });

  it('returns correct Node version for a range with multiple versions', () => {
    const version = spfx.getHighestNodeVersion('^12.13 || ^14.15 || ^16.13');
    assert.strictEqual(version, '16.x');
  });

  it('returns correct Node version when only minor version differ', () => {
    const version = spfx.getHighestNodeVersion('8.1 || 8.2');
    assert.strictEqual(version, '8.x');
  });

  it('returns highest major for disjoint ranges', () => {
    const version = spfx.getHighestNodeVersion('>=16.13.0 <17.0.0 || >=18.17.1 <19.0.0');
    assert.strictEqual(version, '18.x');
  });

  it('returns highest major inclusive upper bound', () => {
    const version = spfx.getHighestNodeVersion('>=14.0.0 <=17.0.0 || >=18.17.1 <=19.0.0');
    assert.strictEqual(version, '19.0.0');
  });

  it('returns exact version for single <= operator', () => {
    const version = spfx.getHighestNodeVersion('<=18.20.4');
    assert.strictEqual(version, '18.20.4');
  });

  it('returns major-1 for exclusive upper bound < operator', () => {
    const version = spfx.getHighestNodeVersion('<17.0.0');
    assert.strictEqual(version, '16.x');
  });

  it('returns correct version for > operator', () => {
    const version = spfx.getHighestNodeVersion('>16.0.0');
    assert.strictEqual(version, '16.x');
  });

  it('returns correct version for exact version without operator', () => {
    const version = spfx.getHighestNodeVersion('16.13.0');
    assert.strictEqual(version, '16.x');
  });

  it('returns highest version when mixing < and <= operators', () => {
    const version = spfx.getHighestNodeVersion('>=14.0.0 <17.0.0 || >=16.0.0 <=18.20.4');
    assert.strictEqual(version, '18.20.4');
  });

  it('throws when range is empty', () => {
    assert.throws(() => spfx.getHighestNodeVersion(''), new Error('Node version range was not provided.'));
  });

  it('throws when range cannot be normalized', () => {
    assert.throws(() => spfx.getHighestNodeVersion('invalid-range'), new Error("Unable to resolve the highest Node version for range 'invalid-range'."));
  });

  it('throws when no valid ranges found', () => {
    assert.throws(() => spfx.getHighestNodeVersion('invalid-string'), new Error("Unable to resolve the highest Node version for range 'invalid-string'."));
  });
});
